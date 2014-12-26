﻿using CRM.Common;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Client;
using Microsoft.Xrm.Sdk.Query;
using System;
using System.Collections.Generic;
using System.Data.SqlClient;
using System.Linq;
using System.Text;

namespace HAT_API_PSADS
{
    class PsadsModel
    {
        //static
        private OrganizationServiceContext xrm = EnvironmentSetting.Xrm;
        private IOrganizationService service = EnvironmentSetting.Service;

        //lookup
        private int new_prdno;
        private int new_account;
        private int new_lotno;
        private int new_salena;

        //date
        private int new_iodate;

        //double
        private int new_dc100;
        private int new_gnum;
        private int new_num;

        //money
        private int new_ntaxamt;
        private int new_price;

        //string
        private int new_unkey;
        private static String[] stringNameArray = { "invoice", "rmk", "saleno", "srno" };
        private int[] stringIntArray = new int[stringNameArray.Length];


        public PsadsModel(SqlDataReader reader)
        {
            try
            {
                //lookup
                new_prdno = reader.GetOrdinal("prdno");
                new_account = reader.GetOrdinal("asno");
                new_lotno = reader.GetOrdinal("lotno");
                new_salena = reader.GetOrdinal("saleno");

                //date
                new_iodate = reader.GetOrdinal("iodate");

                //double
                new_dc100 = reader.GetOrdinal("dc100");
                new_gnum = reader.GetOrdinal("gnum");
                new_num = reader.GetOrdinal("num");

                //money
                new_ntaxamt = reader.GetOrdinal("ntaxamt");
                new_price = reader.GetOrdinal("price");


                //string
                new_unkey = reader.GetOrdinal("unkey");
                for (int i = 0, size = stringNameArray.Length; i < size; i++)
                {
                    stringIntArray[i] = reader.GetOrdinal(stringNameArray[i]);
                }

            }
            catch (Exception ex)
            {
                EnvironmentSetting.ErrorMsg += "搜尋欄位失敗\n";
                EnvironmentSetting.ErrorMsg += ex.Message + "\n";
                EnvironmentSetting.ErrorMsg += ex.Source + "\n";
                EnvironmentSetting.ErrorMsg += ex.StackTrace + "\n";
                EnvironmentSetting.ErrorType = ErrorType.DB;
            }
        }
        public Guid IsCushipExist(SqlDataReader reader)
        {
            try
            {
                return Lookup.RetrieveEntityGuid("new_cuship", reader.GetInt32(new_unkey).ToString(), "new_unkey");
            }
            catch (Exception ex)
            {
                EnvironmentSetting.ErrorMsg += "檢查資料失敗\n";
                EnvironmentSetting.ErrorMsg += ex.Message + "\n";
                EnvironmentSetting.ErrorMsg += ex.Source + "\n";
                EnvironmentSetting.ErrorMsg += ex.StackTrace + "\n";
                EnvironmentSetting.ErrorType = ErrorType.DATASYNCDETAIL;
                return Guid.Empty;
            }
        }
        public TransactionStatus CreateCushipForCRM(SqlDataReader reader)
        {
            try
            {
                Entity entity = new Entity("new_cuship");

                //date
                entity["new_iodate"] = reader.GetDateTime(new_iodate);

                //double
                entity["new_dc100"] = Convert.ToDouble(reader.GetDecimal(new_dc100));
                entity["new_gnum"] = Convert.ToDouble(reader.GetDecimal(new_gnum));
                entity["new_num"] = Convert.ToDouble(reader.GetDecimal(new_num));

                //money
                entity["new_ntaxamt"] = reader.GetDecimal(new_ntaxamt);
                entity["new_price"] = reader.GetDecimal(new_price);

                //string
                entity["new_unkey"] = reader.GetInt32(new_unkey).ToString();
                for (int i = 0, size = stringNameArray.Length; i < size; i++)
                {
                    entity["new_" + stringNameArray[i]] = reader.GetString(stringIntArray[i]).Trim();
                }

                //lookup
                String recordStr;
                Guid recordGuid;

                /// CRM欄位名稱     商品代號    new_grpno
                /// CRM關聯實體     產品        product
                /// CRM關聯欄位     商品代號    productnumber
                /// ERP欄位名稱                 prdno
                /// 
                recordStr = reader.GetString(new_prdno).Trim();
                if (recordStr == "" || recordStr == null)
                    entity["new_prdno"] = null;
                else
                {
                    recordGuid = Lookup.RetrieveEntityGuid("product", recordStr, "productnumber");
                    if (recordGuid == Guid.Empty)
                    {
                        EnvironmentSetting.ErrorMsg = "CRM 查無相符合資料 : \n";
                        EnvironmentSetting.ErrorMsg += "\tCRM實體 : product\n";
                        EnvironmentSetting.ErrorMsg += "\tCRM欄位 : productnumber\n";
                        EnvironmentSetting.ErrorMsg += "\tERP欄位 : prdno\n";
                        Console.WriteLine(EnvironmentSetting.ErrorMsg);
                        return TransactionStatus.Fail;
                    }
                    entity["new_prdno"] = new EntityReference("product", recordGuid);
                }

                /// CRM欄位名稱     客戶          new_account
                /// CRM關聯實體     客戶          account
                /// CRM關聯欄位     客戶編碼      productnumber
                /// ERP欄位名稱                   asno
                /// 
                recordStr = reader.GetString(new_account).Trim();
                if (recordStr == "" || recordStr == null)
                    entity["new_account"] = null;
                else
                {
                    recordGuid = Lookup.RetrieveEntityGuid("account", recordStr, "accountnumber");
                    if (recordGuid == Guid.Empty)
                    {
                        EnvironmentSetting.ErrorMsg = "CRM 查無相符合資料 : \n";
                        EnvironmentSetting.ErrorMsg += "\tCRM實體 : account\n";
                        EnvironmentSetting.ErrorMsg += "\tCRM欄位 : productnumber\n";
                        EnvironmentSetting.ErrorMsg += "\tERP欄位 : asno\n";
                        Console.WriteLine(EnvironmentSetting.ErrorMsg);
                        return TransactionStatus.Fail;
                    }
                    entity["new_account"] = new EntityReference("account", recordGuid);
                }

                /// CRM欄位名稱     批號          new_product_lot
                /// CRM關聯實體     產品批號      new_lotno
                /// CRM關聯欄位     批號          new_lotno
                /// ERP欄位名稱                   lotno
                /// 
                recordStr = reader.GetString(new_lotno).Trim();
                if (recordStr == "" || recordStr == null)
                    entity["new_lotno"] = null;
                else
                {
                    recordGuid = Lookup.RetrieveEntityGuid("product", recordStr, "productnumber");
                    if (recordGuid == Guid.Empty)
                    {
                        EnvironmentSetting.ErrorMsg = "CRM 查無相符合資料 : \n";
                        EnvironmentSetting.ErrorMsg += "\tCRM實體 : new_lotno\n";
                        EnvironmentSetting.ErrorMsg += "\tCRM欄位 : new_lotno\n";
                        EnvironmentSetting.ErrorMsg += "\tERP欄位 : lotno\n";
                        Console.WriteLine(EnvironmentSetting.ErrorMsg);
                        return TransactionStatus.Fail;
                    }
                    entity["new_lotno"] = new EntityReference("new_product_lot", recordGuid);
                }

                /// CRM欄位名稱     業務員     new_salena
                /// CRM關聯實體     使用者     systemuser
                /// CRM關聯欄位     業務代碼   new_saleno
                /// ERP欄位名稱                saleno
                /// 
                recordStr = reader.GetString(new_salena).Trim();
                if (recordStr == "" || recordStr == null)
                    entity["new_salena"] = null;
                else
                {
                    recordGuid = Lookup.RetrieveEntityGuid("systemuser", recordStr, "new_saleno");
                    if (recordGuid == Guid.Empty)
                    {
                        EnvironmentSetting.ErrorMsg = "CRM 查無相符合資料 : \n";
                        EnvironmentSetting.ErrorMsg += "\tCRM實體 : systemuser\n";
                        EnvironmentSetting.ErrorMsg += "\tCRM欄位 : new_saleno\n";
                        EnvironmentSetting.ErrorMsg += "\tERP欄位 : saleno\n";
                        Console.WriteLine(EnvironmentSetting.ErrorMsg);
                        return TransactionStatus.Fail;
                    }
                    entity["new_salena"] = new EntityReference("systemuser", recordGuid);
                }

                try
                {
                    service.Create(entity);
                    return TransactionStatus.Success;
                }
                catch (Exception ex)
                {
                    Console.WriteLine("new_unkey : " + reader["unkey"].ToString().Trim());
                    Console.WriteLine(ex.Message);
                    EnvironmentSetting.ErrorMsg = ex.Message;
                    return TransactionStatus.Fail;
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine("欄位讀取錯誤");
                Console.WriteLine(ex.Message);
                EnvironmentSetting.ErrorMsg = "欄位讀取錯誤\n" + ex.Message;
                return TransactionStatus.Fail;
            }
        }
    }
}
