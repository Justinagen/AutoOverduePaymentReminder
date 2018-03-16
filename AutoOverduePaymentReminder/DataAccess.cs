using System;
using System.Data;
using System.Configuration;
using System.Collections;
using System.Data.SqlClient;

/// <summary>
/// DataAccess 的摘要说明
/// </summary>
public class DataAccess
{

	public DataAccess()
	{
		//
		// TODO: 在此处添加构造函数逻辑
		//
	}
    /// <summary>
    /// 通过Config类从XML文件中得到连接字符串
    /// </summary>
    /// <returns>连接字符串</returns>
    private static string getConnectionString()
    {
        return "server=192.168.1.8;database=pwm;uid=sa;pwd=pwm2252";
    }
    //获取处理器
    public static SqlCommand getCommand()
    {
        
        SqlConnection conn = new SqlConnection(getConnectionString());
        DataSet ds = new DataSet();
        conn.Open();
        SqlCommand cmd = conn.CreateCommand();        
        return cmd;
    }

    //获取带有事物的处理器
    public static SqlCommand getTransactionCommand()
    {
        SqlConnection conn = new SqlConnection(getConnectionString());
        DataSet ds = new DataSet();
        conn.Open();

        SqlTransaction trans = conn.BeginTransaction();
        SqlCommand cmd = conn.CreateCommand();        

        cmd.Transaction = trans;
        return cmd;
    }
    public bool Update(string TableName,Hashtable ht,string Where)
    {
        int Count = 0;
        if (ht.Count <= 0)
        {
            return true;
        }
        String Fields = " ";
        foreach (DictionaryEntry item in ht)
        {
            if (Count != 0)
            {
                Fields += ",";
            }
            Fields += "[" + item.Key.ToString() + "]";
            Fields += "=";
            Fields += "N'"+item.Value.ToString()+"'";
            Count++;
        }
        Fields += " ";

        String SqlString = "Update " + TableName + " Set " + Fields + Where;

        String[] Sqls = { SqlString };
        return ExecuteSQL(Sqls);
    }

    /// <summary>
    /// 公有方法，执行一组Sql语句。
    /// </summary>
    /// <param name="SqlStrings">Sql语句组</param>
    /// <returns>是否成功</returns>
    public bool ExecuteSQL(String[] SqlStrings)
    {
        bool success = true;
        SqlConnection con = new SqlConnection(getConnectionString());
        con.Open();
        SqlCommand cmd = new SqlCommand();
        SqlTransaction trans = con.BeginTransaction();
        cmd.Connection = con;
        cmd.Transaction = trans;
        try
        {
            foreach (String str in SqlStrings)
            {
                cmd.CommandText = str;
                cmd.ExecuteNonQuery();
            }
            trans.Commit();
        }
        catch
        {
            success = false;
            trans.Rollback();
        }
        finally
        {
            con.Close();
        }
        return success;
    }

 //执行数据开始df_jing
    public static bool stExecuteSQL(String[] SqlStrings,string connstring)
    {
        bool success = true;
        SqlConnection con = new SqlConnection(connstring);
        con.Open();
        SqlCommand cmd = new SqlCommand();
        SqlTransaction trans = con.BeginTransaction();
        cmd.Connection = con;
        cmd.Transaction = trans;
        try
        {
            foreach (String str in SqlStrings)
            {
                cmd.CommandText = str;
                cmd.ExecuteNonQuery();
                
            }
            //trans.Rollback();
            trans.Commit();
        }
        catch
        {
            success = false;
            trans.Rollback();
        }
        finally
        {
            con.Close();
        }
        return success;
    }

    /// <summary>
    /// 公有方法，执行Sql语句。
    /// </summary>
    /// <param name="SqlString">Sql语句</param>
    /// <returns>对Update、Insert、Delete为影响到的行数，其他情况为-1</returns>
    public int ExecuteSQL(String SqlString)
    {
        int count = -1;
       
        try
        {
            SqlConnection con = new SqlConnection(getConnectionString());
            con.Open();
            SqlCommand cmd = new SqlCommand(SqlString, con);
            count = cmd.ExecuteNonQuery();
        }
        catch
        {
            count = -1;
        }
       
        return count;
    }


    /// <summary>
    /// 跟据条件，更新表中的数据
    /// </summary>
    /// <param name="dt">要更新的表</param>
    /// <param name="filter">要进行更新的条件</param>
    /// <returns>更新是否成功</returns>
    public static bool SaveDataTable(DataTable dt, string filter)
    {
        try
        {
            //定义一个连接对象
            SqlConnection conn = new SqlConnection(getConnectionString());
            //打开连接
            conn.Open();
            //设置要查询的字段初始内容
            string strFieldList = "*";
            //读取数据表里的字段列表----把要查询的列用表中单独的列名表示出来
            for (int i = 0; i < dt.Columns.Count; i++)
            {
                if (strFieldList.Equals("*"))
                {
                    strFieldList = dt.Columns[i].ColumnName;
                }
                else
                {
                    strFieldList = strFieldList + "," + dt.Columns[i].ColumnName;
                }
            }
            //－－跟据查询语句定义一个适配器对象
            SqlDataAdapter dataAdapter = new SqlDataAdapter("select " + strFieldList + " from " + dt.TableName + filter, conn);
            //－－用命令构建对象，把适配器的所有方法构造出来
            SqlCommandBuilder objCommandBuilder = new SqlCommandBuilder(dataAdapter);
            //－－执行更新操作
            dataAdapter.Update(dt);
            //释放内存空间
            dataAdapter.Dispose();
            //关闭连接
            conn.Close();
            return true;

        }
        catch (SqlException e)
        {
            Console.WriteLine(e.ToString());
            throw new Exception(e.ToString());
        }


    }

    #region 不参考使用事务处理的,使用SQL语句来操作数据库

    /// <summary>
    /// 执行SQL语句，返回单值的结果  -- 用于执行汇总或是结果只返回一行一列的内容
    /// </summary>
    /// <param name="sql">要执行的SQL语句</param>
    /// <returns>返回影响的结果</returns>
    public static int RunSQLReturnIdentity(string sql)
    {
        //定义一个初始的结果
        int intResult = -1;
        try
        {
            //定义一个连接对象
            SqlConnection conn = new SqlConnection(getConnectionString());
            //构造一个命令对象
            SqlCommand cmd = new SqlCommand(sql, conn);
            //打开连接
            cmd.Connection.Open();
            //执行命令，返回单列的结果
            intResult = Convert.ToInt32(cmd.ExecuteScalar());
            conn.Close();
        }
        catch (SqlException e)
        {
            Console.WriteLine(e.ToString());
            throw new Exception(e.ToString() + "\r\n" + sql);
        }
        return intResult;
    }

    /// <summary>
    /// 执行存储过程的方法--此方法用于以文本字符的形式调用存储过程
    /// </summary>
    /// <param name="sql">要执行的SQL语句</param>
    /// <param name="ht">包含所有参数的集合</param>
    /// <returns>执行存储过程的返回结果</returns>
    public static bool RunSQL(string sql, Hashtable ht)
    {
        // 定义参数的名称的数组
        string[] paramsName = new string[ht.Keys.Count];
        // 定义参数值的数组
        object[] paramsValue = new object[ht.Keys.Count];
        // 得到参数名称的遍历器
        IEnumerator ie = ht.Keys.GetEnumerator();
        // 得到所有参数对应的键和值
        for (int i = 0; i < ht.Keys.Count; i++)
        {
            if (ie.MoveNext())
            {
                paramsValue[i] = ht[ie.Current.ToString()];
                paramsName[i] = ie.Current.ToString();
            }
        }
        //定义一个连接对象
        SqlConnection conn = null;
        try
        {
            //创建连接对象
            conn = new SqlConnection(getConnectionString());
            //利用SQL语句，参数名，参数值，连接对象创建命令对象
            SqlCommand cmd = CreateCommand(sql, paramsName, paramsValue, conn);
            //设置是命令类型
            cmd.CommandType = CommandType.Text;
            //打开
            cmd.Connection.Open();
            //执行不返回结果
            cmd.ExecuteNonQuery();
        }
        catch (SqlException e)
        {
            Console.WriteLine(e.ToString());
            throw new Exception(e.ToString() + "\r\n" + sql);
        }
        finally
        {
            if (conn != null) conn.Close();
        }
        return true;
    }

    /// <summary>
    /// 执行SQL，返回DataTable
    /// </summary>
    /// <param name="sql">sql语句</param>
    /// <param name="connString">连接字符串</param>
    /// <returns>返回sql执行后生成的DataTable</returns>
    public static DataTable RunSQLReturnDT(string sql, string connString)
    {
        try
        {
            //连接对象
            SqlConnection conn = new SqlConnection(connString);
            //命令对象
            SqlCommand cmd = new SqlCommand(sql, conn);
            //打开
            cmd.Connection.Open();
            //适配器
            SqlDataAdapter dataAdapter = new SqlDataAdapter();
            //为适配器指定查询命令
            dataAdapter.SelectCommand = cmd;
            //数据表对象
            DataTable dt = new DataTable();
            //填充空的数据表
            dataAdapter.Fill(dt);
            //释放空间 
            dataAdapter.Dispose();
            conn.Close();
            return dt;
        }
        catch (SqlException e)
        {
            Console.WriteLine(e.ToString());
            throw new Exception(e.ToString() + "\r\n" + sql);
        }
    }

    /// 执行SQL，返回DataTable
    /// </summary>
    /// <param name="sql">select语句</param>
    /// <returns>执行SQL，返回DataTable</returns>
    public static DataTable RunSQLReturnDT(string sql)
    {
        try
        {
          
            //连接对象
            SqlConnection conn = new SqlConnection(getConnectionString());
            //命令对象
            SqlCommand cmd = new SqlCommand(sql, conn);
            //打开连接
            cmd.Connection.Open();
            //适配器对象
            SqlDataAdapter dataAdapter = new SqlDataAdapter();
            //设置适配器的查询命令
            dataAdapter.SelectCommand = cmd;
            //数据表对象
            DataTable dt = new DataTable();
            //填充数据表
            dataAdapter.Fill(dt);
            dataAdapter.Dispose();
            conn.Close();
            return dt;
        }
        catch (SqlException e)
        {
            Console.WriteLine(e.ToString());
            throw new Exception(e.ToString() + "\r\n" + sql);
        }
    }

    /// 执行SQL，返回DataTable
    /// </summary>
    /// <param name="sql">select语句</param>
    /// <returns>执行SQL，返回DataTable</returns>
    public static DataTable GetPageData(Int32 pageSize, Int32 pageIndex, string tblName, string condition, out Int32 size)
    {
        Int32 getTotal = 0;
        DataTable dt;
        try
        {
            //连接对象
            SqlConnection conn = new SqlConnection(getConnectionString());
            //命令对象
            SqlCommand cmd = new SqlCommand("[GetPageDataFactoryFun]", conn);
            //打开连接
            cmd.Connection.Open();
            cmd.CommandType = CommandType.StoredProcedure;
            //SqlParameter ppageSize = cmd.Parameters.Add("@pageSize", SqlDbType.Int);
            SqlParameter ppageIndex = cmd.Parameters.Add("@pageIndex", SqlDbType.Int);
            SqlParameter ptblName = cmd.Parameters.Add("@tblName", SqlDbType.VarChar);
            SqlParameter pcondition = cmd.Parameters.Add("@Condition", SqlDbType.VarChar);
            SqlParameter pSize = cmd.Parameters.Add("@Size", SqlDbType.Int);
            //ppageSize.Direction = ParameterDirection.Input;
            ppageIndex.Direction = ParameterDirection.Input;
            ptblName.Direction = ParameterDirection.Input;
            pcondition.Direction = ParameterDirection.Input;
            pSize.Direction = ParameterDirection.Output;

            ppageIndex.Value = pageIndex;
            //ppageSize.Value = pageSize;
            ptblName.Value = tblName;
            pcondition.Value = condition;
            //pSize.Value = size;

            //适配器对象
            SqlDataAdapter dataAdapter = new SqlDataAdapter();
            //设置适配器的查询命令
            dataAdapter.SelectCommand = cmd;
            //数据表对象
            dt = new DataTable();
            //填充数据表
            dataAdapter.Fill(dt);
            size = Convert.ToInt32(pSize.Value.ToString());
            dataAdapter.Dispose();
            conn.Close();
        }
        catch (SqlException e)
        {
            Console.WriteLine(e.ToString());
            throw new Exception(e.ToString());
        }
        return dt;
    }

    /// <summary>
    /// 执行select语句返回DataSet                                               
    /// </summary>
    /// <param name="sql">select语句</param>
    /// <returns>执行select语句返回DataSet</returns>		
    /// <summary>
    public static DataSet RunSQLReturnDS(string sql)
    {
        try
        {
            //连接对象
            SqlConnection conn = new SqlConnection(getConnectionString());
            //命令对象
            SqlCommand cmd = new SqlCommand(sql, conn);
            //打开连接
            cmd.Connection.Open();
            //得到适配器
            SqlDataAdapter dataAdapter = new SqlDataAdapter();
            //设置适配器的查询命令
            dataAdapter.SelectCommand = cmd;
            //数据表对象
            DataTable dt = new DataTable();
            //填充数据表
            dataAdapter.Fill(dt);
            dataAdapter.Dispose();
            conn.Close();
            //把表添加到数据集
            DataSet ds = new DataSet();
            ds.Tables.Add(dt);
            return ds;
        }
        catch (SqlException e)
        {
            Console.WriteLine(e.ToString());
            throw new Exception(e.ToString() + "\r\n" + sql);
        }
    }

    /// <summary>
    /// 执行SQL，返回SqlDataAdapter －－方便在类的外部进行修改操作
    /// </summary>
    /// <param name="sql">select语句</param>
    /// <returns>执行SQL，返回SqlDataAdapter</returns>
    public static SqlDataAdapter RunSQLReturnDA(string sql)
    {
        try
        {
            //连接对象
            SqlConnection conn = new SqlConnection(getConnectionString());
            //命令对象
            SqlCommand cmd = new SqlCommand(sql, conn);
            //打开连接
            cmd.Connection.Open();
            //适配器对象
            SqlDataAdapter dataAdapter = new SqlDataAdapter();
            //设置适配器的查询命令
            dataAdapter.SelectCommand = cmd;
            //利用命令创建对象，把适配器的所有命令创建出来
            SqlCommandBuilder builder = new SqlCommandBuilder(dataAdapter);
            conn.Close();
            return dataAdapter;
        }
        catch (SqlException e)
        {
            Console.WriteLine(e.ToString());
            throw new Exception(e.ToString() + "\r\n" + sql);
        }
    }

    /// <summary>
    /// 执行SQL，返回是否执行成功
    /// </summary>
    /// <param name="sql">sql语句,如:delete,update,insert等</param>
    /// <returns>True:成功,False失败</returns>
    public static bool RunSQL(string sql)
    {
        try
        {
            //连接对象
            SqlConnection conn = new SqlConnection(getConnectionString());
            //命令对象
            SqlCommand cmd = new SqlCommand(sql, conn);
            //打开连接
            cmd.Connection.Open();
            //执行SQL语句
            cmd.ExecuteNonQuery();
            conn.Close();
            return true;
        }
        catch (SqlException e)
        {
            Console.WriteLine(e.ToString());
            throw new Exception(e.ToString() + "\r\n" + sql);
        }

    }

    /// <summary>
    /// 执行SQL，返回是否执行成功
    /// </summary>
    /// <param name="sql">sql语句,如:delete,update,insert等</param>
    /// <returns>True:成功,False失败</returns>
    public static int RunSQLRetuenRows(string sql)
    {
        int i = 0;
        SqlConnection conn = new SqlConnection(getConnectionString());
        if (conn.State == ConnectionState.Closed)
        {
            conn.Open();
        }
        try
        {
            SqlCommand cmd = new SqlCommand(sql, conn);
            i = cmd.ExecuteNonQuery();
            return i;
        }
        catch (Exception ex)
        {
            // strError = "数据检索失败：" + ex.Message;
            return 0;
        }
        finally
        {
            if (conn.State != ConnectionState.Closed)
            {
                conn.Close();
            }
        }
    }

    /// <summary>
    /// 执行SQL，返回Integer-－－影响的行数
    /// </summary>
    /// <param name="sql">sql语句,如:delete,update,insert等</param>
    /// <returns>执行SQL，返回Integer</returns>
    public static int RunSQLReturnInteger(string sql)
    {
        try
        {
            SqlConnection conn = new SqlConnection(getConnectionString());
            SqlCommand cmd = new SqlCommand(sql, conn);
            cmd.Connection.Open();
            SqlDataAdapter dataAdapter = new SqlDataAdapter();
            dataAdapter.SelectCommand = cmd;
            DataTable dt = new DataTable();
            dataAdapter.Fill(dt);
            dataAdapter.Dispose();
            conn.Close();
            int result = 0;
            //判断如果有记录，并且得到的结果不为空
            if (dt.Rows.Count > 0 && !dt.Rows[0][0].ToString().Equals(""))
            {
                result = (int)Int32.Parse(dt.Rows[0][0].ToString());
            }
            return result;
        }
        catch (SqlException e)
        {
            Console.WriteLine(e.ToString());
            throw new Exception(e.ToString() + "\r\n" + sql);
        }

    }
 

    #region 未完成
    public static bool RunSQL(string sql, string[] paramsName, object[] paramsValue)
    {
        SqlConnection conn = null;
        try
        {
            conn = new SqlConnection(getConnectionString());
            SqlCommand cmd = CreateCommand(sql, paramsName, paramsValue, conn);
            cmd.CommandType = CommandType.Text;
            cmd.Connection.Open();
            cmd.ExecuteNonQuery();
        }
        catch (SqlException e)
        {
            Console.WriteLine(e.ToString());
            throw new Exception(e.ToString() + "\r\n" + sql);
        }
        finally
        {
            if (conn != null) conn.Close();
        }
        return true;
    }


    //public static string RunSQLReturnID(string sql,string [] paramsName,object [] paramsValue)
    //{
    //    string result = "";
    //    SqlConnection conn = null;
    //    try
    //    {
    //        conn = new SqlConnection(getConnectionString());
    //        SqlCommand cmd = CreateCommand(sql,paramsName,paramsValue,conn);
    //        cmd.CommandType=CommandType.Text;
    //        cmd.Connection.Open();
    //        object obj = cmd.ExecuteScalar();
    //        if(obj!=System.DBNull.Value)
    //        {
    //            result = obj.ToString();
    //        }
    //    }
    //    catch(SqlException e) 
    //    {
    //        Console.WriteLine(e.ToString());
    //        throw new Exception(e.ToString() + "\r\n" + sql);
    //    }
    //    finally
    //    {
    //        if(conn!=null)conn.Close();
    //    }
    //    return result;		
    //}

    private static SqlCommand CreateCommand(string sql, String[] paramsName, Object[] paramsValue, SqlConnection conn)
    {
        SqlCommand cmd = null;

        //			if (paramsName != null) 
        //			{
        //				for(int i=0;i<paramsName.Length;i++)
        //				{
        //					sql = Lang.replaceFirst(sql,"@",paramsName[i]);
        //				}
        //				
        //				cmd = new SqlCommand(sql, conn);
        //				
        //				for(int i=0;i<paramsName.Length;i++)
        //				{
        //					sql = Lang.replaceFirst(sql,"@",paramsName[i]);
        //					cmd.Parameters.Add(paramsName[i],paramsValue[i]);
        //				}
        //			}
        return cmd;
    }

    #endregion
    #endregion

}
