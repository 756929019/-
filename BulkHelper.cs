using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Data;
using CYQ.Data;
using System.Data.SqlClient;
using System.ComponentModel;
using Business.Sql;

namespace Water.Utilites
{
    public class BulkHelper
    {
        //事件
        public event System.Data.SqlClient.SqlRowsCopiedEventHandler BulkInsertEvent;
        public event DoWorkEventHandler DoWorkEvent;
        public event ProgressChangedEventHandler ProgressChangedEvent;
        public event RunWorkerCompletedEventHandler RunWorkerCompletedEvent;
        public string ErrorText { get; set; }
        /// <summary>
        /// 后台线程
        /// </summary>
        private BackgroundWorker bkWorker = null;
        public bool CancellationPending {
            get {
                if (bkWorker == null)
                {
                    return false;
                }
                else
                {
                    return bkWorker.CancellationPending;
                }
            }
        }
        public void InitBgWork()
        {
            if (bkWorker == null)
            {
                bkWorker = new BackgroundWorker();
                bkWorker.WorkerReportsProgress = true;
                bkWorker.WorkerSupportsCancellation = true;
                if (DoWorkEvent != null)
                {
                    bkWorker.DoWork += new DoWorkEventHandler(DoWorkEvent);
                }
                if (ProgressChangedEvent != null)
                {
                    bkWorker.ProgressChanged += new ProgressChangedEventHandler(ProgressChangedEvent);
                }
                if (RunWorkerCompletedEvent != null)
                {
                    bkWorker.RunWorkerCompleted += new RunWorkerCompletedEventHandler(RunWorkerCompletedEvent);
                }
            }
        }
        
        public void StartWork()
        {
            bkWorker.RunWorkerAsync();
        }

        public void StartWork(object argument)
        {
            bkWorker.RunWorkerAsync(argument);
        }

        public void CancelWork()
        {
            bkWorker.CancelAsync();
        }

        public void ReportProgress(int rp)
        {
            bkWorker.ReportProgress(rp);
        }

        /// <summary>
        /// 大批量插入数据
        /// </summary>
        /// <param name="tabList">批量插入表列表</param>
        /// <param name="sqlList">更新数据sql(在批量插入表之后)</param>
        /// <param name="prepareSqlList">更新数据sql(在批量插入表之前)</param>
        /// <param name="MyConStr">连接字符串</param>
        /// <returns></returns>
        public int TransferData(List<BulkTable> tabList, List<SqlModel> sqlList, List<SqlModel> prepareSqlList=null,string MyConStr = "")
        {
            try
            {
                //if (tabList == null || tabList.Count == 0)
                //{
                //    return -1;
                //}
                //获取连接字符串
                string conStr = "";
                if (string.IsNullOrEmpty(MyConStr))
                {
                    using (MAction action = new MAction(TableNames.DICTCONTENT))
                    {
                        conStr = action.ConnectionString;
                    }
                }
                else
                {
                    conStr = MyConStr;
                }
                
                //一个连接对象
                SqlConnection con = new SqlConnection(conStr);
                //打开连接
                con.Open();
                //事务
                SqlTransaction Trans = con.BeginTransaction();
                try
                {
                    if (prepareSqlList != null && prepareSqlList.Count > 0)
                    {
                        SqlCommand command;
                        int rint = 0;
                        foreach (SqlModel sm in prepareSqlList)
                        {
                            command = new SqlCommand(sm.sql, con);
                            command.CommandTimeout = 0;
                            command.Transaction = Trans;
                            rint = command.ExecuteNonQuery();
                            if (sm.isNew)
                            {
                                //影响条数必须大于0
                                if (rint < 1)
                                {
                                    Trans.Rollback();
                                    ErrorText = sm.errorText;
                                    con.Close();
                                    return -2;
                                }
                            }
                            if (sm.resultCount > 0)
                            {
                                //有预计影响条数，不匹配回滚
                                if (rint != sm.resultCount)
                                {
                                    Trans.Rollback();
                                    ErrorText = sm.errorText;
                                    con.Close();
                                    return -3;
                                }
                            }
                            //更新sql且没有预计影响条数
                            if (rint < 0)
                            {
                                Trans.Rollback();
                                ErrorText = sm.errorText;
                                con.Close();
                                return -4;
                            }
                        }
                    }

                    if (tabList != null && tabList.Count > 0)
                    {
                        //批量插入数据
                        using (System.Data.SqlClient.SqlBulkCopy bcp = new System.Data.SqlClient.SqlBulkCopy(con, SqlBulkCopyOptions.CheckConstraints, Trans))
                        {
                            if (BulkInsertEvent != null)
                                bcp.SqlRowsCopied += new System.Data.SqlClient.SqlRowsCopiedEventHandler(BulkInsertEvent);
                            bcp.BatchSize = 100;              //每次传输的行数
                            bcp.NotifyAfter = 100;            //进度提示的行数
                            foreach (BulkTable bt in tabList)
                            {
                                bcp.DestinationTableName = bt.tableName; //目标表
                                bcp.WriteToServer(bt.dataTable, DataRowState.Added);
                            }
                        }
                    }
                    
                    if (sqlList != null && sqlList.Count > 0)
                    {
                        SqlCommand command;
                        int rint = 0;
                        foreach (SqlModel sm in sqlList)
                        {
                            command = new SqlCommand(sm.sql, con);
                            command.CommandTimeout = 0;
                            command.Transaction = Trans;
                            rint = command.ExecuteNonQuery();
                            if (sm.isNew)
                            {
                                //影响条数必须大于0
                                if (rint < 1)
                                {
                                    Trans.Rollback();
                                    ErrorText = sm.errorText;
                                    con.Close();
                                    return -2;
                                }
                            }
                            if (sm.resultCount > 0)
                            {
                                //有预计影响条数，不匹配回滚
                                if (rint != sm.resultCount)
                                {
                                    Trans.Rollback();
                                    ErrorText = sm.errorText;
                                    con.Close();
                                    return -3;
                                }
                            }
                            //更新sql且没有预计影响条数
                            if (rint < 0)
                            {
                                Trans.Rollback();
                                ErrorText = sm.errorText;
                                con.Close();
                                return -4;
                            }
                        }
                    }
                    Trans.Commit();
                }
                catch (Exception ex)
                {
                    Trans.Rollback();
                    ErrorText = ex.Message;
                    con.Close();
                    return -5;
                }
                finally
                {
                    con.Close();
                }
                return 1;
            }
            catch (Exception ex1)
            {
                ErrorText = ex1.Message;
                return -6;
            }
        }

        public class BulkTable
        {
            public string tableName { get; set; }
            public DataTable dataTable { get; set; }
        }

        public class SqlModel
        {
            public bool isNew { get; set; }
            public string sql { get; set; }
            public string errorText { get; set; }
            public int resultCount { get; set; }//预计影响条数，0不做比较
        }
    }
}
