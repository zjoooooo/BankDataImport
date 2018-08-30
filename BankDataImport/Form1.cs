using System;
using System.Collections.Generic;
using System.Data;
using System.Data.OleDb;
using System.Data.SqlClient;
using System.IO;
using System.Threading;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;

namespace BankDataImport
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }
        Dictionary<string, string> CP_Cepass_Pair = new Dictionary<string, string>();
        Dictionary<string, string> CP_Mechant_Pair = new Dictionary<string, string>();
        private void Form1_Load(object sender, EventArgs e)
        {
            initContainer();
            Thread thr = new Thread(() => letsdoit());
            thr.Start();
        }

        private void initContainer()
        {
            CP_Cepass_Pair.Clear();
            CP_Mechant_Pair.Clear();
            //string constr = "Data Source=172.16.1.6;uid=sa;pwd=NVT123;database=SECUREPARK";
            string constr = "Data Source=172.16.1.6;uid=secure;pwd=weishenme;database=SECUERPARK";
            string cmd = @"SELECT * FROM [dbo].[CarParkCodeDetails] WHERE Carpark12digitNo!='' ORDER BY Carpark12digitNo;
                           SELECT * FROM CarParkCodeDetails WHERE Carpark4digitNo!='' and CarparkTypeID=3 ORDER BY Carpark4digitNo;";
            DataSet ds = null;
            try
            {
                ds = SqlHelper.ExecuteDataset(constr, CommandType.Text, cmd);

            }
            catch (SqlException e)
            {
                LogClass.WriteLog("Fail To Get Container List!");
            }

            string GetReapt = null;
            foreach (DataRow dr in ds.Tables[0].Rows)
            {
                string carpark = dr["CarParkID"].ToString();
                string merchant_ID = dr["Carpark12digitNo"].ToString();
                if (GetReapt == merchant_ID)
                {
                    continue;
                }
                CP_Mechant_Pair.Add(merchant_ID, carpark);
                GetReapt = merchant_ID;
                //LogClass.WriteLog(merchant_ID);
            }


            string GetLtaReapt = null;
            foreach (DataRow dr in ds.Tables[1].Rows)
            {
                string carpark = dr["CarParkID"].ToString();
                string cepass_ID = dr["Carpark4digitNo"].ToString();
                if (GetLtaReapt == cepass_ID)
                {
                    continue;
                }
                CP_Cepass_Pair.Add(cepass_ID, carpark);
                GetLtaReapt = cepass_ID;
                //LogClass.WriteLog(cepass_ID);
            }

        }

        private void letsdoit()
        {
            DirectoryInfo TheFolder = new DirectoryInfo(Application.StartupPath);

            foreach (FileInfo NextFile in TheFolder.GetFiles())
            {
                if ((NextFile.Extension.Equals(".csv")) || (NextFile.Extension.Equals(".xls")) || (NextFile.Extension.Equals(".xlsx")))
                {
                    LetsRead(NextFile, NextFile.Name.Split('.')[0].ToUpper());

                }
            }
            Application.Exit();
        }

        private void LetsRead(FileInfo file, string type)
        {
            string DBString = @"Provider=Microsoft.ACE.OLEDB.12.0;Data Source =" + file.FullName + ";Extended Properties=Excel 12.0";
            OleDbConnection con = new OleDbConnection(DBString);
            con.Open();
            DataTable datatable = con.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, null);
            //获取表单的名字
            String sheet = datatable.Rows[0][2].ToString().Trim();
            // string sheet = "Sheet1";
            OleDbDataAdapter ole = new OleDbDataAdapter("select * from [" + sheet + "]", con);
            DataSet ds = new DataSet();
            ole.Fill(ds);
            con.Close();
            LogClass.WriteLog($"====== {type} Report ======");

            if (type.Equals("NETS"))
            {
                LogClass.WriteLog($"retailer_id,CASHCARD_PURCHASE,CASHCARD_transaction,Credit,Debit");
            }
            else if (type.Equals("LTA"))
            {
                LogClass.WriteLog("Carpark_Owner_UENID,Carpark_Owner_ID,Carpark_ID,Card_Manager,LTA_Process_Date,Transaction_Date,Total_Processed_Count,Total_Processed_Amount");
            }
            string constr = "Data Source=172.16.1.89;uid=secure;pwd=weishenme;database=NetsSettlementAudit";
            foreach (DataRow col in ds.Tables[0].Rows)
            {
                switch (type)
                {
                    case "NETS":
                        string Retailer_Id = col[0].ToString();
                        string[] str_array = Retailer_Id.Split('(');
                        string pre_retailer = null;
                        if (str_array.Length > 1)
                        {
                             pre_retailer = str_array[1];
                        }
                        else
                        {
                            continue;
                        }

                        Retailer_Id = pre_retailer.Substring(0, pre_retailer.Length - 1);
                        //LogClass.WriteLog(Retailer_Id);
                        string Cashcard_Purchase = col[1].ToString();
                        string Cashcard_Transaction = col[2].ToString();
                        string Credit = col[3].ToString();
                        string Debit = col[4].ToString();
                        string Transaction_dt = col[5].ToString();
                        //string Merchant_ID =;
                        LogClass.WriteLog($"{Retailer_Id},{Cashcard_Purchase},{Cashcard_Transaction},{Credit},{Debit},{Transaction_dt}");
                        //update nets db.
                        if (!CP_Mechant_Pair.TryGetValue(Retailer_Id, out string carpark_nets))
                        {
                            carpark_nets = "NuknowCP";
                        }

                        string cmd_nets = @"Insert INTO [NETS_Details](Retailer_Id,Cashcard_Purchase,Cashcard_Transaction,Credit,Debit,CarParkName,Transaction_dt,Update_dt)
                                                         VALUES(@Retailer_Id,@Cashcard_Purchase,@Cashcard_Transaction,@Credit,@Debit,@CarParkName,@Transaction_dt,getdate())";
                        SqlParameter[] para_nets = new SqlParameter[]
                        {
                            new SqlParameter("@Retailer_Id",Retailer_Id),
                            new SqlParameter("@Cashcard_Purchase",Cashcard_Purchase),
                            new SqlParameter("@Cashcard_Transaction",Cashcard_Transaction),
                            new SqlParameter("@Credit",Credit),
                            new SqlParameter("@Debit",Debit),
                            new SqlParameter("@CarParkName",carpark_nets),
                            new SqlParameter("@Transaction_dt",Transaction_dt)
                        };

                        try
                        {
                            SqlHelper.ExecuteNonQuery(constr, CommandType.Text, cmd_nets, para_nets);
                        }
                        catch (SqlException netse)
                        {
                            LogClass.WriteLog($"Fail To Insert NETS Data Into Db.{netse.ToString()}");
                        }

                        break;
                    case "LTA":
                        string Carpark_Owner_UENID = col[0].ToString();
                        string Carpark_Owner_ID = col[1].ToString();
                        string Carpark_ID = col[2].ToString();
                        string Card_Manager = col[3].ToString();
                        string LTA_Process_Date = col[4].ToString();
                        LTA_Process_Date = (Convert.ToDateTime(LTA_Process_Date)).ToString("yyyy-MM-dd HH:mm:ss");
                        string Transaction_Date = col[5].ToString();
                        Transaction_Date = (Convert.ToDateTime(Transaction_Date)).ToString("yyyy-MM-dd HH:mm:ss");
                        string Total_Processed_Count = col[6].ToString();
                        string Total_Processed_Amount = col[7].ToString();
                        LogClass.WriteLog($"{Carpark_Owner_UENID},{Carpark_Owner_ID},{Carpark_ID},{Card_Manager},{LTA_Process_Date},{Transaction_Date},{Total_Processed_Count},{Total_Processed_Amount}");
                        //update lta db.

                        if (!CP_Cepass_Pair.TryGetValue(Carpark_ID, out string carpark_lta))
                        {
                            carpark_lta = "NuknowCP";
                        }

                        string cmd_lta = @"Insert INTO [LTA_Details](CarparkOwnerUENID,CarparkOwnerID,CarparkID,CardManager,LTAProcessDate,TransactionDate,TotalProcessedCount,TotalProcessedAmount,CarParkName,Update_dt)
                                                         VALUES(@Carpark_Owner_UENID,@Carpark_Owner_ID,@Carpark_ID,@Card_Manager,@LTA_Process_Date,@Transaction_Date,@Total_Processed_Count,@Total_Processed_Amount,@CarParkName,getdate())";
                        SqlParameter[] para_lta = new SqlParameter[]
                        {
                            new SqlParameter("@Carpark_Owner_UENID",Carpark_Owner_UENID),
                            new SqlParameter("@Carpark_Owner_ID",Carpark_Owner_ID),
                            new SqlParameter("@Carpark_ID",Carpark_ID),
                            new SqlParameter("@Card_Manager",Card_Manager),
                            new SqlParameter("@LTA_Process_Date",LTA_Process_Date),
                            new SqlParameter("@Transaction_Date",Transaction_Date),
                            new SqlParameter("@Total_Processed_Count",Total_Processed_Count),
                            new SqlParameter("@Total_Processed_Amount",Total_Processed_Amount),
                            new SqlParameter("@CarParkName",carpark_lta)
                        };

                        try
                        {

                            //SqlHelper.ExecuteNonQuery(constr, CommandType.Text, cmd_lta, para_lta);
                        }
                        catch (SqlException ltae)
                        {
                            LogClass.WriteLog($"Fail To Insert LTA Data Into Db.{ltae.ToString()}");
                        }
                        break;
                }
            }
        }

    }
}
