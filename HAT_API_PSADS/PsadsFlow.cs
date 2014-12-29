using CRM.Common;
using CRM.Model;
using NLog;
using System;
using System.Collections.Generic;
using System.Data.SqlClient;
using System.Linq;
using System.Text;

namespace HAT_API_PSADS
{
    class PsadsFlow
    {
        private static Logger _logger = LogManager.GetCurrentClassLogger();
        public void DoPsadsFlow()
        {
            DataSyncModel dataSync = new DataSyncModel();
            //connection CRM
            EnvironmentSetting.LoadSetting();
            if (EnvironmentSetting.ErrorType == ErrorType.None)
            {
                //create dataSync
                dataSync.CreateDataSyncForCRM("客戶出貨明細");
                if (EnvironmentSetting.ErrorType == ErrorType.None)
                {
                    //connection DB
                    EnvironmentSetting.LoadDB();
                    if (EnvironmentSetting.ErrorType == ErrorType.None)
                    {
                        //get information from DB
                        EnvironmentSetting.GetList("select * from dbo.v_psads where salena like '%趙%'");
                        if (EnvironmentSetting.ErrorType == ErrorType.None)
                        {
                            //get reader
                            SqlDataReader reader = EnvironmentSetting.Reader;

                            //search field index
                            PsadsModel psadsmodel = new PsadsModel(reader);
                            if (EnvironmentSetting.ErrorType == ErrorType.None)
                            {
                                Console.WriteLine("連線成功!!");
                                Console.WriteLine("開始執行...");

                                if (reader.HasRows)
                                {
                                    int success = 0;
                                    int fail = 0;
                                    int partially = 0;
                                    while (reader.Read())
                                    {
                                        //判斷CRM是否有資料
                                        Guid existCushipId = psadsmodel.IsCushipExist(EnvironmentSetting.Reader);
                                        if (EnvironmentSetting.ErrorType == ErrorType.None)
                                        {
                                            TransactionStatus transactionStatus = TransactionStatus.None;
                                            TransactionType transactionType = TransactionType.Insert;
                                            if (existCushipId == Guid.Empty)
                                            {
                                                //create
                                                transactionStatus = psadsmodel.CreateCushipForCRM(EnvironmentSetting.Reader);
                                            }

                                            if (EnvironmentSetting.ErrorType == ErrorType.None)
                                            {
                                                //create datasyncdetail
                                                switch (transactionStatus)
                                                {
                                                    case TransactionStatus.Success:
                                                        success++;
                                                        break;
                                                    case TransactionStatus.Fail:
                                                        //新增、更新資料有錯誤 則新增一筆detail
                                                        dataSync.CreateDataSyncDetailForCRM(reader["unkey"].ToString().Trim(), reader["unkey"].ToString().Trim(), transactionType, transactionStatus);
                                                        fail++;
                                                        break;
                                                    default:
                                                        break;
                                                }

                                                //新增detail錯誤 則結束
                                                if (EnvironmentSetting.ErrorType != ErrorType.None)
                                                {
                                                    _logger.Info(EnvironmentSetting.ErrorMsg);
                                                    break;
                                                }
                                            }
                                        }
                                    }
                                    //更新DataSync 成功、失敗、完成時間
                                    dataSync.UpdateDataSyncForCRM(success, fail, partially);
                                }
                                else
                                {
                                    Console.WriteLine("沒有資料");
                                    EnvironmentSetting.ErrorMsg += "ERP沒有資料\n";
                                    EnvironmentSetting.ErrorType = ErrorType.DB;
                                }
                            }
                        }
                    }
                }
            }
            switch (EnvironmentSetting.ErrorType)
            {
                case ErrorType.None:
                    break;

                case ErrorType.INI:
                case ErrorType.CRM:
                case ErrorType.DATASYNC:
                    _logger.Info(EnvironmentSetting.ErrorMsg);
                    break;

                case ErrorType.DB:
                case ErrorType.DATASYNCDETAIL:
                    dataSync.UpdateDataSyncWithErrorForCRM(EnvironmentSetting.ErrorMsg);
                    break;

                default:
                    break;
            }
            Console.WriteLine("執行完畢...");
            Console.ReadLine();
        }

    }
}
