using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Collections;
using System.IO;
using System.Windows.Forms;
using System.Data;
using System.Data.SqlClient;

namespace WarrantAssistant
{
    
    public class SafeQueue : IDisposable
    {
        private Queue dataList = new Queue();
        private object lockObj = new object();
        private bool isUsing = true;

        public void Enqueue(object data)
        {
            if (!isUsing) { return; }
            lock (lockObj)
            {
                dataList.Enqueue(data);
            }
        }
        public object Dequeue()
        {
            if (dataList.Count > 0)
            {
                object data;
                lock (lockObj)
                {
                    data = dataList.Dequeue();
                }
                return data;
            }
            else
                return null;
        }
        public int Count
        {
            get
            {
                return dataList.Count;
            }
        }
        #region IDisposable成員
        public void Dispose()
        {
            isUsing = false;
            dataList.Clear();
        }
        #endregion
    }

    public class TXTFileWriter : IDisposable
    {
        public string fileName = "";
        private bool isInitialOK = false;
        private object lockObj = new object();
        private FileStream fileStream;
        private StreamWriter streamWriter;

        public TXTFileWriter(string fileName)
        {
            this.fileName = fileName;
            InitialFileWriter();
        }
        private void InitialFileWriter()
        {
            try
            {
                if (File.Exists(fileName))
                {
                    File.Delete(fileName);
                }

                fileStream = new FileStream(fileName, FileMode.Append);
                /*
                using (FileStream fs = new FileStream(fileName, FileMode.Append))
                {
                    using (StreamWriter sw = new StreamWriter(fs,enc)
                    {
                        
                    }
                }
                 * */
                streamWriter = new StreamWriter(fileStream, Encoding.GetEncoding("Big5"));
                isInitialOK = true;
            }
            catch (Exception ex)
            {
                MessageBox.Show("[" + DateTime.Now.ToString("HH:mm:ss.fff") + "][TXTFileWriter_InitialFileWriter][" + ex.Message + "] " + fileName);
            }
        }
        public void WriteFile(string writeString)
        {
            if (!isInitialOK) { return; }
            if (streamWriter == null) { return; }
            lock (lockObj)
            {
                streamWriter.WriteLine(writeString);
            }
        }
        #region IDisposable成員
        public void Dispose()
        {
            try
            {
                isInitialOK = false;
                lock (lockObj)
                {
                    if (streamWriter != null)
                    {
                        streamWriter.Close();
                        streamWriter.Dispose();
                    }
                    if (fileStream != null)
                    {
                        fileStream.Close();
                        fileStream.Dispose();
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }
        #endregion
    }

    public class SQLCommandHelper : IDisposable
    {
        public string connString = "";
        public string commandString = "";
        public List<SqlParameter> parameterList;

        //private bool isInitialOK = false;
        private SqlCommand cmd;

        public SQLCommandHelper(string connString, string commandString, List<SqlParameter> parameterList)
        {
            this.connString = connString;
            this.commandString = commandString;
            this.parameterList = parameterList;
            InitialCommand();
        }
        private void InitialCommand()
        {
            try
            {
                SqlConnection conn = new SqlConnection {
                    ConnectionString = connString
                };
                cmd = new SqlCommand {
                    Connection = conn,
                    CommandText = commandString
                };
                foreach (SqlParameter parameter in parameterList)
                    cmd.Parameters.Add(parameter);

                //isInitialOK = true;
            }
            catch (Exception ex)
            {
                MessageBox.Show("[" + DateTime.Now.ToString("HH:mm:ss.fff") + "][SQLCommandHelper_InitialCommand][" + ex.Message + "]");
            }
        }
        public void SetParameterValue(string parameterName, object value)
        {
            try
            {
                if (value != null)
                    cmd.Parameters[parameterName].Value = value;
                else
                    cmd.Parameters[parameterName].Value = DBNull.Value;
            }
            catch (Exception ex)
            {
                //GlobalVar.errProcess.Add(1, "[SqlCommandHelper_SetParameterValue][" + ex.Message + "][" + ex.StackTrace + "]");
            }

        }
        public void ExecuteCommand()
        {
            try
            {
                if (cmd.Connection.State != ConnectionState.Open)
                    cmd.Connection.Open();
                cmd.ExecuteNonQuery();
                cmd.Connection.Close();
            }
            catch (Exception ex)
            {
                MessageBox.Show("[" + DateTime.Now.ToString("HH:mm:ss.fff") + "][SQLCommandHelper][" + ex.Message + "]");
            }
        }
        #region IDisposable成員
        public void Dispose()
        {
            if (cmd != null && cmd.Connection != null && cmd.Connection.State != ConnectionState.Closed)
            {
                cmd.Connection.Close();
                cmd.Connection.Dispose();
                cmd.Dispose();
            }
        }
        #endregion
    }
}
