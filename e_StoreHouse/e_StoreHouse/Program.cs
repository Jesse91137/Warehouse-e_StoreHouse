using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.Data;
using System.Data.SqlClient;
using System.IO;


namespace e_StoreHouse
{
    class Program
    {
        /// <summary>
        /// Db connection string
        /// </summary>
        private static readonly String connStr = ConfigurationManager.AppSettings["CNN_TEXT"].ToString();

        /// <summary>
        /// 來源-正式
        /// </summary>
        static string sourceP = @"\\192.168.4.11\33 倉庫備料單\07. 收發組\成倉\成倉SAP_Excel\";

        /* 複製檔案到 D:\成品倉\每日匯入Excel工具\Temp\ */
        //static string destP = @"D:\成品倉\每日匯入Excel工具\Temp\";

        /// <summary>
        /// 取得執行檔所在目錄，並組合 Temp 資料夾路徑
        /// 從[\\192.168.4.11\33 倉庫備料單\07. 收發組\成倉\成倉SAP_Excel\]複製檔案到[執行檔目錄\Temp\]
        /// </summary>
        static string destP = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Temp") + Path.DirectorySeparatorChar;



        //static List<string> li = new List<string>() {
        //    @"\\192.168.4.11\33 倉庫備料單\07. 收發組\成倉\成倉SAP_Excel\ZRSD19.xlsx",
        //    @"\\192.168.4.11\33 倉庫備料單\07. 收發組\成倉\成倉SAP_Excel\ZRSD14P.xlsx",
        //    @"\\192.168.4.11\33 倉庫備料單\07. 收發組\成倉\成倉SAP_Excel\ZRSD13.xlsx",
        //    @"\\192.168.4.11\33 倉庫備料單\07. 收發組\成倉\成倉SAP_Excel\KQ30.xlsx",
        //    @"\\192.168.4.11\33 倉庫備料單\07. 收發組\成倉\成倉SAP_Excel\KF10.xlsx" };
        static List<string> li = new List<string>() { };
        static void Main(string[] args)
        {
            Copy_Paste();
            ToDo();
        }

        /// <summary>
        /// 複製檔案到 D:\成品倉\每日匯入Excel工具\Temp\
        /// </summary>
        public static void Copy_Paste()
        {
            DirectoryInfo di = new DirectoryInfo(sourceP);

            foreach (var item in di.GetFiles())
            {
                // 取得副檔名，並轉成小寫
                string ext = item.Extension.ToLower();
                // 如果副檔名是 .xls 或 .xlsx

                if (ext == ".xls" || ext == ".xlsx")
                {
                    System.IO.File.Copy(sourceP + item.Name, destP + item.Name, true);
                    li.Add(destP + item.Name);
                }
            }
        }

        /// <summary>
        /// 執行讀取Excel寫入DB
        /// </summary>
        private static void ToDo()
        {
            #region 執行前刪除前一天資料
            string _19d = "delete from E_ZRSD19";
            ExecueNonQuery(_19d, CommandType.Text, null);

            string _14d = "delete from E_ZRSD14P";
            ExecueNonQuery(_14d, CommandType.Text, null);
            string _13d = "delete from E_ZRSD13";
            ExecueNonQuery(_13d, CommandType.Text, null);
            string _SCd = "delete from E_StoreHouseStock_SC";
            ExecueNonQuery(_SCd, CommandType.Text, null);
            string _10d = "delete from E_Kf10";
            ExecueNonQuery(_10d, CommandType.Text, null);
            string _30d = "delete from E_KQ30";
            ExecueNonQuery(_30d, CommandType.Text, null);
            //string _QCd = "delete from E_QC";
            //ExecueNonQuery(_QCd, CommandType.Text, null);
            #endregion

            DataTable dt = new DataTable();
            try
            {
                for (int i = 0; i < li.Count; i++)
                {
                    string xlsFileName = li[i];
                    string[] strli = xlsFileName.Split('\\');
                    string keyWord = strli[4];
                    dt = LoadExcelAsDataTable(xlsFileName, keyWord);

                    foreach (DataRow row in dt.Rows)
                    {
                        switch (keyWord)
                        {
                            #region ZRSD19
                            case string a when a.Contains("ZRSD19"):
                                /*  Excel 欄位
                                 *   工單、客戶訂單、客戶工單單號、料號、成品規格
                                 * 、訂單數量、已出貨量、未出貨量、借出數量、工單開立數量
                                 * 、工單入庫數量、預交日、生產注意事項_1、生產注意事項_2、業務單號
                                 * 、業單項次
                                 */
                                string insSql19 = @"INSERT INTO E_ZRSD19 ([wono],[order_cust],[wono_cust],[item],[spec_product]
                                                                ,[order_quantity],[shipped_quantity],[unshipped_quantity],[borrow_count]
                                                                ,[wono_openCount],[wono_inStoreCount],[due_date],[prod_notes],[prod_notes2])
                                                                VALUES(@z191,@z192,@z193,@z194,@z195
                                                                        ,@z196,@z197,@z198,@z199,@z1910
                                                                        ,@z1911,@z1912,@z1913,@z1914)";
                                SqlParameter[] parm19 = new SqlParameter[]
                                {
                                                new SqlParameter("z191",row[0].ToString()),// 工單
                                                new SqlParameter("z192",row[1].ToString()),// 客戶訂單
                                                new SqlParameter("z193",row[2].ToString()),// 客戶工單單號
                                                new SqlParameter("z194",row[3].ToString()),// 料號
                                                new SqlParameter("z195",row[4].ToString()),// 成品規格
                                                new SqlParameter("z196",row[5].ToString()),// 訂單數量
                                                new SqlParameter("z197",row[6].ToString()),// 已出貨量
                                                new SqlParameter("z198",row[7].ToString()),// 未出貨量
                                                new SqlParameter("z199",row[8].ToString()),// 借出數量
                                                new SqlParameter("z1910",row[9].ToString()),// 工單開立數量
                                                new SqlParameter("z1911",row[10].ToString()),// 工單入庫數量
                                                new SqlParameter("z1912",row[11].ToString()),// 預交日
                                                new SqlParameter("z1913",row[12].ToString()),// 生產注意事項_1
                                                new SqlParameter("z1914",row[13].ToString()),// 生產注意事項_2
                                };
                                try
                                {
                                    ExecueNonQuery(insSql19, CommandType.Text, parm19);
                                }
                                catch (Exception _19e)
                                {
                                    Console.WriteLine("_19e" + _19e.Message);
                                    Console.ReadKey();
                                }
                                #endregion
                                break;
                            #region ZRSD14P
                            case string b when b.Contains("ZRSD14P"):
                                /* Excel 欄位
                                 *   工單、預交日、料號、客戶工單單號、成品規格
                                 * 、訂單數量、己出貨量、未出貨量、借出數量、生產注意事項_1
                                 */
                                string insSql14 = @"INSERT INTO E_ZRSD14P([wono],[due_date],[item],[wono_cust],[spec_product]
                                                                    ,[order_quantity],[shipped_quantity],[unshipped_quantity],[borrow_count],[prod_notes])
                                                                    VALUES(@z141,@z142,@z143,@z144,@z145
                                                                            ,@z146,@z147,@z148,@z149,@z1410)";
                                SqlParameter[] parm14 = new SqlParameter[]
                                {
                                                new SqlParameter("z141",row[0].ToString()),// 工單
                                                new SqlParameter("z142",row[1].ToString()),// 預交日
                                                new SqlParameter("z143",row[2].ToString()),// 料號
                                                new SqlParameter("z144",row[3].ToString()),// 客戶工單單號
                                                new SqlParameter("z145",row[4].ToString()),// 成品規格
                                                new SqlParameter("z146",row[5].ToString()),// 訂單數量
                                                new SqlParameter("z147",row[6].ToString()),// 己出貨量
                                                new SqlParameter("z148",row[7].ToString()),// 未出貨量
                                                new SqlParameter("z149",row[8].ToString()),// 借出數量
                                                new SqlParameter("z1410",row[9].ToString()),// 生產注意事項_1
                                };
                                try
                                {
                                    ExecueNonQuery(insSql14, CommandType.Text, parm14);
                                }
                                catch (Exception _14e)
                                {
                                    Console.WriteLine("_14e" + _14e.Message);
                                    Console.ReadKey();
                                }
                                #endregion
                                break;
                            #region ZRSD13
                            case string c when c.Contains("ZRSD13"):
                                /*  Excel 欄位
                                 *   Plan GI Date、Sales Order、料號、規格說明、客戶工單號碼
                                 * 、數量、Order Item、出貨單備註
                                 */
                                if (!string.IsNullOrEmpty(row[1].ToString()))
                                {
                                    string insSql13 = @"INSERT INTO E_ZRSD13([gi_date],[sales_order],[item],[spec_desc],[cust_wono]
                                                                            ,[quantity],[order_item],[ship_mark])
                                                                      VALUES(@z131,@z132,@z133,@z134,@z135
                                                                            ,@z136,@z137,@z138)";
                                    SqlParameter[] parm13 = new SqlParameter[]
                                    {
                                                new SqlParameter("z131",row[0].ToString()),// Plan GI Date
                                                new SqlParameter("z132",row[1].ToString()),// Sales Order
                                                new SqlParameter("z133",row[2].ToString()),// 料號
                                                new SqlParameter("z134",row[3].ToString()),// 規格說明
                                                new SqlParameter("z135",row[4].ToString()),// 客戶工單號碼
                                                new SqlParameter("z136",row[5].ToString()),// 數量
                                                new SqlParameter("z137",row[6].ToString()),// Order Item
                                                new SqlParameter("z138",row[7].ToString()),// 出貨單備註
                                    };
                                    try
                                    {
                                        ExecueNonQuery(insSql13, CommandType.Text, parm13);
                                    }
                                    catch (Exception _13e)
                                    {
                                        Console.WriteLine("_13e" + _13e.Message);
                                        Console.ReadKey();
                                    }
                                }
                                #endregion
                                break;
                            #region KF10
                            case string d when d.Contains("KF10"):
                                /* Excel 欄位
                                 *   物料、物料說明、Specification、工廠、SLoc
                                 * 、S、SL、批次、BUn、未限制
                                 * 、移轉中庫存、品質檢驗中、限制使用的庫存、已凍結、退貨
                                 * 、銷售文件
                                 */
                                string insSql10 = @"INSERT INTO E_Kf10([Material],[Material_Desc],[Specification],[Plnt],[SLoc]
                                                                       ,[S],[SL],[Batch],[BUn],[Unrestricted]
                                                                       ,[Stock_in_Transfer],[I_Q_I],[Restricted_Use],[Blocked],[RE]
                                                                       ,[Docu])
                                                                 VALUES(@k101,@k102,@k103,@k104,@k105,
                                                                        @k106,@k107,@k108,@k109,@k1010,
                                                                        @k1011,@k1012,@k1013,@k1014,@k1015,
                                                                        @k1016)";
                                SqlParameter[] parm10 = new SqlParameter[]
                                {
                                                new SqlParameter("k101",row[0].ToString()),// 物料
                                                new SqlParameter("k102",row[1].ToString()),// 物料說明
                                                new SqlParameter("k103",row[2].ToString()),// Specification
                                                new SqlParameter("k104",row[3].ToString()),// 工廠
                                                new SqlParameter("k105",row[4].ToString()),// SLoc
                                                new SqlParameter("k106",row[5].ToString()),// S
                                                new SqlParameter("k107",row[6].ToString()),// SL
                                                new SqlParameter("k108",row[7].ToString()),// 批次
                                                new SqlParameter("k109",row[8].ToString()),// BUn
                                                new SqlParameter("k1010",row[9].ToString()),// 未限制
                                                new SqlParameter("k1011",row[10].ToString()),// 移轉中庫存
                                                new SqlParameter("k1012",row[11].ToString()),// 品質檢驗中
                                                new SqlParameter("k1013",row[12].ToString()),// 限制使用的庫存
                                                new SqlParameter("k1014",row[13].ToString()),// 已凍結
                                                new SqlParameter("k1015",row[14].ToString()),// 退貨
                                                new SqlParameter("k1016",row[15].ToString()),// 銷售文件
                                };
                                try
                                {
                                    ExecueNonQuery(insSql10, CommandType.Text, parm10);
                                }
                                catch (Exception _10e)
                                {
                                    Console.WriteLine("_10e" + _10e.Message);
                                    Console.ReadKey();
                                }
                                #endregion
                                break;
                            #region KQ30
                            case string e when e.Contains("KQ30"):
                                /* Excel 欄位
                                 *   物料、物料說明、Specification、工廠、SLoc
                                 * 、S、SL、BUn、未限制、移轉中庫存
                                 * 、品質檢驗中、限制使用的庫存、已凍結、退貨、銷售文件
                                 */
                                string insSql30 = @"INSERT INTO E_KQ30([Material],[Material_Desc],[Specification],[Plnt],[SLoc]
                                                                    ,[S],[SL],[Batch],[BUn],[Unrestricted]
                                                                    ,[Stock_in_Transfer],[I_Q_I],[Restricted_Use],[Blocked],[RE]
                                                                    ,[Docu])
                                                                    VALUES(@k301,@k302,@k303,@k304,@k305,
                                                                           @k306,@k307,@k308,@k309,@k3010,
                                                                           @k3011,@k3012,@k3013,@k3014,@k3015,@k3016)";
                                SqlParameter[] parm30 = new SqlParameter[]
                                {
                                                new SqlParameter("k301",row[0].ToString()),// 物料
                                                new SqlParameter("k302",row[1].ToString()),// 物料說明
                                                new SqlParameter("k303",row[2].ToString()),// Specification
                                                new SqlParameter("k304",row[3].ToString()),// 工廠
                                                new SqlParameter("k305",row[4].ToString()),// SLoc
                                                new SqlParameter("k306",row[5].ToString()),// S
                                                new SqlParameter("k307",row[6].ToString()),// SL
                                                new SqlParameter("k308",row[7].ToString()),// BUn
                                                new SqlParameter("k309",row[8].ToString()),// 未限制
                                                new SqlParameter("k3010",row[9].ToString()),// 移轉中庫存
                                                new SqlParameter("k3011",row[10].ToString()),// 品質檢驗中
                                                new SqlParameter("k3012",row[11].ToString()),// 限制使用的庫存
                                                new SqlParameter("k3013",row[12].ToString()),// 已凍結
                                                new SqlParameter("k3014",row[13].ToString()),// 退貨
                                                new SqlParameter("k3015",row[14].ToString()),// 銷售文件
                                                new SqlParameter("k3016",row[15].ToString()),
                                };
                                try
                                {
                                    ExecueNonQuery(insSql30, CommandType.Text, parm30);
                                }
                                catch (Exception _30e)
                                {
                                    Console.WriteLine("_30e" + _30e.Message);
                                    Console.ReadKey();
                                }
                                #endregion
                                break;
                                #region _QC
                                //case string f when f.Contains("QC"):
                                /* QC、檢驗單號、iclass、出貨序號、生產序號
                                 * 、不良代碼、描述、執行狀態、IsSample、Is PASS
                                 * 、檢驗結果、工作人員、委外廠、工單號碼、GUID
                                 * 、不良代碼、生產日期
                                 */

                                //    string insSqlQC = @"INSERT INTO E_QC([QCinspect],[iclass],[ship_serialno],[product_serialno],[error_code]
                                //                                                                                ,[describe],[status],[issample],[ispass],[inspect_result],[op_user]
                                //                                                                                ,[outsourcing],[wono],[GUID],[error2],[product_date])
                                //                                        VALUES(@c1,@c2,@c3,@c4,@c5,@c6,@c7,
                                //                                                        @c8,@c9,@c10,@c11,@c12,@c13,@c14,@c15,@c16)";
                                //    SqlParameter[] parmQC = new SqlParameter[]
                                //    {
                                //                    new SqlParameter("c1",row[0].ToString()),
                                //                    new SqlParameter("c2",row[1].ToString()),
                                //                    new SqlParameter("c3",row[2].ToString()),
                                //                    new SqlParameter("c4",row[3].ToString()),
                                //                    new SqlParameter("c5",row[4].ToString()),
                                //                    new SqlParameter("c6",row[5].ToString()),
                                //                    new SqlParameter("c7",row[6].ToString()),
                                //                    new SqlParameter("c8",row[7].ToString()),
                                //                    new SqlParameter("c9",row[8].ToString()),
                                //                    new SqlParameter("c10",row[9].ToString()),
                                //                    new SqlParameter("c11",row[10].ToString()),
                                //                    new SqlParameter("c12",row[11].ToString()),
                                //                    new SqlParameter("c13",row[12].ToString()),
                                //                    new SqlParameter("c14",row[13].ToString()),
                                //                    new SqlParameter("c15",row[14].ToString()),
                                //                    new SqlParameter("c16",row[15].ToString()),
                                //    };
                                //    try
                                //    {
                                //        ExecueNonQuery(insSqlQC, CommandType.Text, parmQC);
                                //    }
                                //    catch (Exception _QCe)
                                //    {
                                //        Console.WriteLine("_QCe" + _QCe.Message);
                                //        Console.ReadKey();
                                //    }

                                //    break;
                                #endregion
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
                Console.ReadKey();
            }
            //寫入預計出貨
            try
            {
                string insSqlSc = @"insert into E_StoreHouseStock_SC 
                                    (exp_shipdate,sales_order,cust_wono,order_item,eng_sr,exp_shipquantity,sap_mark,UserId,IsApproved)
                                    (select gi_date,sales_order,cust_wono,order_item,item,quantity,ship_mark,'00000','N' from E_ZRSD13)";
                ExecueNonQuery(insSqlSc, CommandType.Text, null);
            }
            catch (Exception _sc)
            {
                Console.WriteLine("_sc:" + _sc.Message);
                Console.ReadKey();
            }

            //更新倉庫庫存的Kf10 . Kq30資料
            try
            {
                string ukf10 = @"update a set a.kf10 = b.Unrestricted 
                                from E_StoreHouseStock a
                                join E_Kf10 b on substring(a.wono,1,7) = b.docu
                                where substring(a.wono,1,7) in 
                                (select docu from E_Kf10 where len(docu)>1 group by docu having count(docu)<2) and quantity>0 and del_flag is null ";
                ExecueNonQuery(ukf10, CommandType.Text, null);
            }
            catch (Exception e_ulf10)
            {
                Console.WriteLine(e_ulf10.Message);
                Console.ReadKey();
            }

            try
            {
                string ukq30 = @"update a set a.kq30=c.Unrestricted
                                from E_StoreHouseStock a
                                join E_KQ30 c on substring(a.wono,1,7) = c.docu
                                where substring(a.wono,1,7) in 
                                (select docu from E_KQ30 where len(docu)>1 group by docu having count(docu)<2) and quantity>0 and del_flag is null";
                ExecueNonQuery(ukq30, CommandType.Text, null);
            }
            catch (Exception e_ukq30)
            {
                Console.WriteLine(e_ukq30.Message);
                Console.ReadKey();
            }


            Console.WriteLine("\n\n\n\n" + "寫入完畢,按任意建關閉!!");
        }

        /// <summary>
        /// 讀取Excel檔案，轉成DataTable
        /// </summary>
        /// <param name="xlsFilename"></param>
        /// <param name="key"></param>
        /// <returns>DataTable</returns>
        public static DataTable LoadExcelAsDataTable(String xlsFilename, string key)
        {
            FileInfo fi = new FileInfo(xlsFilename);
            using (FileStream fstream = new FileStream(fi.FullName, FileMode.Open))
            {
                IWorkbook wb;
                if (fi.Extension == ".xlsx")
                    wb = new XSSFWorkbook(fstream); // excel2007
                else
                    wb = new HSSFWorkbook(fstream); // excel97

                // 只取第一個sheet。
                ISheet sheet = wb.GetSheetAt(0);

                // target
                DataTable table = new DataTable();
                // 由第一列取標題做為欄位名稱
                IRow headerRow = null;
                int cellCount = 0;
                int iFirstRowNum = 0;
                switch (key)
                {
                    case string a when a.Contains("ZRSD19"):
                    case string b when b.Contains("ZRSD14P"):
                    case string c when c.Contains("ZRSD13"):
                        headerRow = sheet.GetRow(5);
                        iFirstRowNum = 7;
                        break;
                    case string d when d.Contains("KF10"):
                    case string e when e.Contains("KQ30"):
                        headerRow = sheet.GetRow(1);
                        iFirstRowNum = 3;
                        break;
                    case string f when f.Contains("QC"):
                        headerRow = sheet.GetRow(0);
                        iFirstRowNum = 1;
                        break;
                }
                cellCount = headerRow.LastCellNum; // 取欄位數
                try
                {
                    for (int i = headerRow.FirstCellNum; i < cellCount; i++)
                    {
                        //table.Columns.Add(new DataColumn(headerRow.GetCell(i).StringCellValue, typeof(double)));
                        //table.Columns.Add(new DataColumn(headerRow.GetCell(i).StringCellValue));
                        string columnName = headerRow.GetCell(i)?.StringCellValue;
                        //if (!string.IsNullOrEmpty(columnName))
                        //{
                        //      table.Columns.Add(new DataColumn(columnName));
                        //}
                        if (table.Columns.Contains(columnName))
                        {
                            table.Columns.Add(new DataColumn(columnName + i));
                        }
                        else
                        {
                            table.Columns.Add(new DataColumn(columnName));
                        }

                        //table.Columns.Add(new DataColumn("Cell_" + i));
                    }
                }
                catch (Exception exx)
                {
                    Console.WriteLine("exx:" + exx.Message);
                    Console.ReadKey();
                    throw;
                }

                try
                {
                    // 略過第零列(標題列)，一直處理至最後一列
                    for (int i = iFirstRowNum; i <= sheet.LastRowNum; i++)
                    {
                        IRow row = sheet.GetRow(i);
                        if (row == null) continue;

                        DataRow dataRow = table.NewRow();

                        //依先前取得的欄位數逐一設定欄位內容

                        for (int j = row.FirstCellNum; j < cellCount; j++)
                        {
                            ICell cell = row.GetCell(j);
                            if (cell != null)
                            {
                                //如要針對不同型別做個別處理，可善用.CellType判斷型別
                                //再用.StringCellValue, .DateCellValue, .NumericCellValue...取值

                                switch (cell.CellType)
                                {
                                    case CellType.Numeric:
                                        if (DateUtil.IsCellDateFormatted(cell))
                                        {
                                            // 如果單元格格式是日期格式，就把數值轉換為日期
                                            dataRow[j] = cell.DateCellValue.ToString("yyyy-MM-dd");
                                            // 處理日期數據
                                        }
                                        else
                                        {
                                            // 如果不是日期格式，就當做數字處理
                                            dataRow[j] = cell.NumericCellValue;
                                            // 處理數字數據
                                        }
                                        //dataRow[j] = cell.NumericCellValue;
                                        break;
                                    default: // String
                                             //此處只簡單轉成字串
                                        dataRow[j] = cell.StringCellValue;
                                        break;
                                }
                            }
                        }

                        table.Rows.Add(dataRow);
                    }
                }
                catch (Exception excel)
                {
                    Console.WriteLine("excel:" + excel.Message);
                    Console.ReadKey();
                    throw;
                }
                // success
                return table;
            }
        }

        /// <summary>
        /// 執行sql語句
        /// </summary>
        /// <param name="sql"></param>
        /// <param name="cmdType"></param>
        /// <param name="pms"></param>
        /// <returns></returns>
        public static int ExecueNonQuery(string sql, CommandType cmdType, params SqlParameter[] pms)
        {
            using (SqlConnection con = new SqlConnection(connStr))
            {
                using (SqlCommand cmd = new SqlCommand(sql, con))
                {
                    //設置目前執行的是「存儲過程? 還是帶參數的sql 語句?」
                    cmd.CommandType = cmdType;
                    if (pms != null)
                    {
                        cmd.Parameters.AddRange(pms);
                    }

                    con.Open();
                    return cmd.ExecuteNonQuery();
                }
            }
        }
    }
}
