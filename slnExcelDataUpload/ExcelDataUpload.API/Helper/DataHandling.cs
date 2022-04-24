using Microsoft.Extensions.Configuration;
using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;

namespace ExcelDataUpload.API.Helper
{
    public class DataHandling
    {
        private readonly string Dbconnectionstring; 

        public readonly IConfiguration  _configuration;

        public DataHandling(IConfiguration configuration)
        {
            _configuration = configuration ?? throw new ArgumentNullException(nameof(configuration));
            Dbconnectionstring = _configuration.GetValue<string>("DatabaseSettings:ConnectionString");     
        }

        public bool ProcessDataAdotoSQL(DataSet srcDataSet, string dbDestinationTblNm,List<string> dbDestTableList, bool blncolumnMappingActionSts=true)
        {
            Boolean datatransferSts = false;
            try
            {
                if(srcDataSet != null && srcDataSet.Tables.Count>0)
                {
                    foreach(var table in srcDataSet.Tables)
                    {
                        using(SqlConnection objSqlConn=new SqlConnection(Dbconnectionstring))
                        {
                            objSqlConn.Open();
                            if (dbDestTableList !=null && dbDestTableList.Count>0)
                            {
                                //Write Code if any validation on Table Data Required or Data Require to delete first.
                            }

                            using (SqlBulkCopy objSqlBulkCopy = new SqlBulkCopy(objSqlConn))
                            {
                                objSqlBulkCopy.BatchSize = 100000;
                                objSqlBulkCopy.BulkCopyTimeout = 1000;
                                objSqlBulkCopy.DestinationTableName = dbDestinationTblNm;
                                objSqlBulkCopy.ColumnMappings.Clear();

                                if(blncolumnMappingActionSts)
                                {
                                    foreach (DataColumn objcolumn in srcDataSet.Tables[0].Columns)
                                    {
                                        objSqlBulkCopy.ColumnMappings.Add(objcolumn.ColumnName, objcolumn.ColumnName.ToUpper());
                                    }
                                }

                                objSqlBulkCopy.WriteToServer(srcDataSet.Tables[0].CreateDataReader());
                                datatransferSts = true;
                            }
                        }
                    }
                }

            }
            catch (Exception)
            {
                datatransferSts = false;
                throw;
            }
            finally
            {
                if (dbDestTableList != null)
                { dbDestTableList.Clear();}


                if (srcDataSet != null)
                {
                    srcDataSet=null;
                }

            }

            return datatransferSts;
;
        }
    }
}
