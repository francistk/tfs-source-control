using System;
using System.Collections;
using System.Configuration;
using System.Data;
using System.Data.OleDb;

namespace DataAccess
{
  /// <summary>
  /// AccessHelper 数据逻辑访问帮助类
  /// update 2004-10-07
  /// </summary>
  public class AccessHelper
  {
    /// <summary>
    /// 数据库连接串
    /// </summary>  
    //public static string CONN_STRING_ACCESS = "Provider=Microsoft.Jet.OleDb.4.0;Data Source=" + AppDomain.CurrentDomain.BaseDirectory.ToString() + ConfigurationSettings.AppSettings["CONNSTR_Access"];
	  public static string CONN_STRING_ACCESS = GetConnStr();

    //public static readonly string CONN_STRING_SQL = ConfigurationSettings.AppSettings["CONNSTR_SQL"];

    // 存储参数的哈希表
    private static Hashtable parmCache = Hashtable.Synchronized(new Hashtable());

	public AccessHelper()
  	{

	}

	public static string GetConnStr()
	  {
		  string conn = "";
		  //先转化成小写,以免区分大小写
          //string strTemp = ConfigurationManager.AppSettings["Webdiy_DBType"].ToLower();
          //if (strTemp == "access")
          conn = "Provider=Microsoft.Jet.OleDb.4.0;Data Source=" + AppDomain.CurrentDomain.BaseDirectory.ToString() + ConfigurationManager.AppSettings["MsnRobotDbPath"];
          //else if (strTemp == "sqlserver")
          //    conn = "Provider=sqloledb;" + ConfigurationManager.AppSettings["CONNSTR_SQL"];
          //    //throw new Exception("个人版不支持sql数据库!");
          //else
          //    throw new Exception("配置文件错误!");

		  return conn;
	  }

    /// <summary>
    /// 执行一条没有返回数据结果集的SQL语句或存储过程
    /// </summary>
    /// <param name="connString">有效的数据库连接串</param>
    /// <param name="cmdType">SQL命令类型：SQL语句或存储过程</param>
    /// <param name="cmdText">命令字符串</param>
    /// <param name="cmdParms">执行需要的参数数组</param>
    /// <returns>影响的行数</returns>
    public static int ExecuteNonQuery(string connString, CommandType cmdType, string cmdText, params OleDbParameter[] cmdParms)
    {
      OleDbCommand cmd = new OleDbCommand();

      // 在 using 范围外自动销毁 conn 对象
      using (OleDbConnection conn = new OleDbConnection(connString))
      {
        PrepareCommand(cmd, conn, null, cmdType, cmdText, cmdParms);
        int val = cmd.ExecuteNonQuery();
        cmd.Parameters.Clear();
        return val;
      }
    }

    /// <summary>
    /// 执行一条没有返回数据结果集的SQL语句或存储过程
    /// </summary>
    /// <param name="conn">数据库连接对象</param>
    /// <param name="cmdType">SQL命令类型：SQL语句或存储过程</param>
    /// <param name="cmdText">命令字符串</param>
    /// <param name="cmdParms">执行需要的参数数组</param>
    /// <returns>影响的行数</returns>
    public static int ExecuteNonQuery(OleDbConnection conn, CommandType cmdType, string cmdText, params OleDbParameter[] cmdParms)
    {
      OleDbCommand cmd = new OleDbCommand();

      PrepareCommand(cmd, conn, null, cmdType, cmdText, cmdParms);
      int val = cmd.ExecuteNonQuery();
      cmd.Parameters.Clear();
      return val;
    }

    /// <summary>
    /// 执行一条没有返回数据结果集的SQL语句或存储过程
    /// </summary>
    /// <param name="trans">存在的事务</param>
    /// <param name="cmdType">SQL命令类型：SQL语句或存储过程</param>
    /// <param name="cmdText">命令字符串</param>
    /// <param name="cmdParms">执行需要的参数数组</param>
    /// <returns>影响的行数</returns>
    public static int ExecuteNonQuery(OleDbTransaction trans, CommandType cmdType, string cmdText, params OleDbParameter[] cmdParms)
    {
      OleDbCommand cmd = new OleDbCommand();

      PrepareCommand(cmd, trans.Connection, trans, cmdType, cmdText, cmdParms);
      int val = cmd.ExecuteNonQuery();
      cmd.Parameters.Clear();
      return val;
    }

    /// <summary>
    /// 执行一条SQL命令，返回 OleDbDataReader
    /// </summary>
    /// <param name="connString">数据库连接串</param>
    /// <param name="cmdType">SQL命令类型：SQL语句或存储过程</param>
    /// <param name="cmdText">命令字符串</param>
    /// <param name="cmdParms">执行需要的参数数组</param>
    /// <returns>影响的行数</returns>
    public static OleDbDataReader ExecuteReader(string connString, CommandType cmdType, string cmdText, params OleDbParameter[] cmdParms)
    {
      OleDbCommand cmd = new OleDbCommand();
      OleDbConnection conn = new OleDbConnection(connString);

      // 这里用 try/catch 机制主要是考虑到，当SQL命令执行异常时又没有建立 OleDbDataReader 对象，
      // 自然也就无法使用 CommandBehavior.CloseConnection 来关闭数据库连接对象 conn
      // 因此，这里用 try/catch 机制，当异常时进行关闭数据库连接对象 conn
      try
      {
        PrepareCommand(cmd, conn, null, cmdType, cmdText, cmdParms);
        OleDbDataReader rdr = cmd.ExecuteReader(CommandBehavior.CloseConnection);
        cmd.Parameters.Clear();
        return rdr;
      }
      catch
      {
        conn.Close();
        throw;
      }
    }


    /// <summary>
    /// 执行一条 SQL 命令，返回一个单元格的值
    /// </summary>
    /// <param name="connString">数据库连接串</param>
    /// <param name="cmdType">SQL命令类型：SQL语句或存储过程</param>
    /// <param name="cmdText">命令字符串</param>
    /// <param name="cmdParms">执行需要的参数数组</param>
    /// <returns>object</returns>
    public static object ExecuteScalar(string connString, CommandType cmdType, string cmdText, params OleDbParameter[] cmdParms)
    {
      OleDbCommand cmd = new OleDbCommand();

      using (OleDbConnection conn = new OleDbConnection(connString))
      {
        PrepareCommand(cmd, conn, null, cmdType, cmdText, cmdParms);
        object val = cmd.ExecuteScalar();
        cmd.Parameters.Clear();
        return val;
      }
    }

    /// <summary>
    /// 执行一条 SQL 命令，返回一个单元格的值
    /// </summary>
    /// <param name="conn">数据库连接对象</param>
    /// <param name="cmdType">SQL命令类型：SQL语句或存储过程</param>
    /// <param name="cmdText">命令字符串</param>
    /// <param name="cmdParms">执行需要的参数数组</param>
    /// <returns>object</returns>
    public static object ExecuteScalar(OleDbConnection conn, CommandType cmdType, string cmdText, params OleDbParameter[] cmdParms)
    {
      OleDbCommand cmd = new OleDbCommand();

      PrepareCommand(cmd, conn, null, cmdType, cmdText, cmdParms);
      object val = cmd.ExecuteScalar();
      cmd.Parameters.Clear();
      return val;
    }

    /// <summary>
    /// 执行一条 SQL 命令，返回取得的数据集
    /// </summary>
    /// <param name="connString">数据库连接串</param>
    /// <param name="cmdType">SQL命令类型：SQL语句或存储过程</param>
    /// <param name="cmdText">命令字符串</param>
    /// <param name="cmdParms">执行需要的参数数组</param>
    /// <returns>结果数据集</returns>
    public static DataSet ExecuteDataset(string connString, CommandType cmdType, string cmdText, params OleDbParameter[] cmdParms)
    {
      OleDbCommand cmd = new OleDbCommand();
      OleDbConnection conn = new OleDbConnection(connString);
      DataSet ds = new DataSet();

      try
      {
        PrepareCommand(cmd, conn, null, cmdType, cmdText, cmdParms);
        OleDbDataAdapter da = new OleDbDataAdapter();
        da.SelectCommand = cmd;
        da.Fill(ds);
        return ds;
      }
      catch
      {
        throw;
      }
      finally
      {
        conn.Close();
      }
    }

    /// <summary>
    /// 执行一条 SQL 命令，返回取得的数据集
    /// </summary>
    /// <param name="conn">数据库连接对象</param>
    /// <param name="cmdType">SQL命令类型：SQL语句或存储过程</param>
    /// <param name="cmdText">命令字符串</param>
    /// <param name="cmdParms">执行需要的参数数组</param>
    /// <returns>结果数据集</returns>
    public static DataSet ExecuteDataset(OleDbConnection conn, CommandType cmdType, string cmdText, params OleDbParameter[] cmdParms)
    {
      OleDbCommand cmd = new OleDbCommand();
      DataSet ds = new DataSet();

      try
      {
        PrepareCommand(cmd, conn, null, cmdType, cmdText, cmdParms);
        OleDbDataAdapter da = new OleDbDataAdapter();
        da.SelectCommand = cmd;
        da.Fill(ds);
        return ds;
      }
      catch
      {
        throw;
      }
      finally
      {
        conn.Close();
      }
    }

    /// <summary>
    /// 将参数数组存入哈希表
    /// </summary>
    /// <param name="cacheKey">Key</param>
    /// <param name="cmdParms">要存入的数组</param>
    public static void CacheParameters(string cacheKey, params OleDbParameter[] cmdParms)
    {
      parmCache[cacheKey] = cmdParms;
    }

    /// <summary>
    /// 取得哈希表里的参数
    /// </summary>
    /// <param name="cacheKey">Key</param>
    /// <returns>参数数组值</returns>
    public static OleDbParameter[] GetCachedParameters(string cacheKey)
    {
      OleDbParameter[] cachedParms = (OleDbParameter[]) parmCache[cacheKey];

      if (cachedParms == null)
        return null;

      OleDbParameter[] clonedParms = new OleDbParameter[cachedParms.Length];

      for (int i = 0, j = cachedParms.Length; i < j; i++)
        clonedParms[i] = (OleDbParameter) ((ICloneable) cachedParms[i]).Clone();

      return clonedParms;
    }

    /// <summary>
    /// 对参数数组进行准备
    /// </summary>
    /// <param name="cmd">OleDbCommand 对象</param>
    /// <param name="conn">OleDbConnection 对象</param>
    /// <param name="trans">OleDbTransaction 对象</param>
    /// <param name="cmdType">SQL命令类型：SQL语句或存储过程</param>
    /// <param name="cmdText">命令字符串</param>
    /// <param name="cmdParms">参数数组</param>
    private static void PrepareCommand(OleDbCommand cmd, OleDbConnection conn, OleDbTransaction trans, CommandType cmdType, string cmdText, OleDbParameter[] cmdParms)
    {
      if (conn.State != ConnectionState.Open)
        conn.Open();

      cmd.Connection = conn;
      cmd.CommandText = cmdText;

      if (trans != null)
        cmd.Transaction = trans;

      cmd.CommandType = cmdType;

      if (cmdParms != null)
      {
        foreach (OleDbParameter parm in cmdParms)
          cmd.Parameters.Add(parm);
      }
    }

    #region 新增加的重载方法
    //**************************新增的部分开始***********************************
    /// <summary>
    /// 执行一条 SQL 命令，返回DataSet
    /// </summary>
    /// <param name="cmdText"></param>
    /// <returns></returns>
    public static DataSet ExecuteDataset(string cmdText)
    {
      OleDbCommand cmd = new OleDbCommand();
      OleDbConnection conn = new OleDbConnection(AccessHelper.CONN_STRING_ACCESS);
      DataSet ds = new DataSet();

      try
      {
        if (conn.State != ConnectionState.Open)
          conn.Open();

        cmd.Connection = conn;
        cmd.CommandText = cmdText;
        cmd.CommandType = CommandType.Text;

        OleDbDataAdapter da = new OleDbDataAdapter();
        da.SelectCommand = cmd;
        da.Fill(ds);
        return ds;
      }
      catch
      {
        throw;
      }
      finally
      {
        conn.Close();
      }
    }

    /// <summary>
    /// 执行一条 SQL 命令，返回DataTable
    /// </summary>
    /// <param name="cmdText"></param>
    /// <returns></returns>
    public static DataTable ExecuteDataTable(string cmdText)
    {
      DataSet ds = ExecuteDataset(cmdText);
      if (ds.Tables.Count > 0)
        return ds.Tables[0];
      else
        return null;
    }

    /// <summary>
    /// 执行一条 SQL 命令，返回DataRow
    /// </summary>
    /// <param name="cmdText"></param>
    /// <returns></returns>
    public static DataRow ExecuteDataRow(string cmdText)
    {
      DataSet ds = ExecuteDataset(cmdText);
      if (ds.Tables.Count > 0 && ds.Tables[0].Rows.Count > 0)
        return ds.Tables[0].Rows[0];
      else
        return null;
    }

    /// <summary>
    /// 执行一条 SQL 命令，返回所影响的行数
    /// </summary>
    /// <param name="cmdText"></param>
    /// <returns></returns>
    public static int ExecuteNonQuery(string cmdText)
    {
      OleDbCommand cmd = new OleDbCommand();

      // 在 using 范围外自动销毁 conn 对象
      using (OleDbConnection conn = new OleDbConnection(AccessHelper.CONN_STRING_ACCESS))
      {
        PrepareCommand(cmd, conn, null, CommandType.Text, cmdText, null);
        int val = cmd.ExecuteNonQuery();
        cmd.Parameters.Clear();
        return val;
      }
    }

    /// <summary>
    /// 执行一条 SQL 命令，返回第一行,第一列
    /// </summary>
    /// <param name="cmdText"></param>
    /// <returns></returns>
    public static object ExecuteScalar(string cmdText)
    {
      OleDbCommand cmd = new OleDbCommand();

      using (OleDbConnection conn = new OleDbConnection(AccessHelper.CONN_STRING_ACCESS))
      {
        PrepareCommand(cmd, conn, null, CommandType.Text, cmdText, null);
        object val = cmd.ExecuteScalar();
        cmd.Parameters.Clear();
        return val;
      }
    }

      public static OleDbDataReader ExecuteReader(string cmdText)
      {
          OleDbCommand cmd = new OleDbCommand();
          OleDbConnection conn = new OleDbConnection(AccessHelper.CONN_STRING_ACCESS);
          try
          {
              PrepareCommand(cmd, conn, null, CommandType.Text, cmdText, null);
              OleDbDataReader rdr = cmd.ExecuteReader(CommandBehavior.CloseConnection);
              cmd.Parameters.Clear();
              return rdr;
          }
          catch
          {
              conn.Close();
              throw;
          }
      }
   //**************************新增的部分结束***********************************
      #endregion
  }
}