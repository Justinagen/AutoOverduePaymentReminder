using System;
using System.Collections;
using System.Configuration;
using System.Data;
using System.Data.SqlClient;
using System.Net;
using System.Net.Mail;
using System.Text;
/// <summary>
/// Utils 的摘要说明
/// </summary>
public abstract class Utils
{    
    
    public static readonly string loginStr = "<script language='javascript' type='text/javascript'>top.opener=null;top.open('','_self');top.open('login.aspx','','toolbar=yes,location=yes,status=yes,menubar=yes,scrollbars=yes,resizable=yes,fullscreen=0');top.close();</script>";
    public static readonly string errorStr = "<script language='javascript' type='text/javascript'>window.open('ErrorPage.aspx','_self');</script>";

    public static  DataTable executeQueryT(string  sql, string connStr )
    {
        string strError = ""; 
        DataTable dt = new DataTable("Table1");
        SqlConnection conn = new SqlConnection(connStr);
        if (conn.State == ConnectionState.Closed)
        {
            conn.Open();
        }
        try
        {
            SqlDataAdapter mySqlDataAdapter = new SqlDataAdapter();
            SqlCommand selectCommand = new SqlCommand(sql, conn);
            selectCommand.CommandTimeout = 9000;
            selectCommand.CommandType = CommandType.Text;
            mySqlDataAdapter.SelectCommand = selectCommand;
            mySqlDataAdapter.Fill(dt);
            return dt;

        }
        catch (Exception ex)
        {
            strError = "数据检索失败：" + ex.Message;
            return null; 
        }
        finally
        {
            if (conn.State != ConnectionState.Closed)
            {
                conn.Close();
            } 
        }
        
    }

    public  static int executeUpdate(string sql,string connStr) {
        string strError = ""; 
        int i = 0;
        SqlConnection conn = new SqlConnection(connStr);
        if (conn.State == ConnectionState.Closed)
        {
            conn.Open();
        }
        try
        {
            SqlCommand cmd = new SqlCommand(sql,conn);
            cmd.CommandTimeout = 900;
            i = cmd.ExecuteNonQuery();
            return i;
        }
        catch (Exception ex)
        {
            strError = "数据检索失败：" + ex.Message;
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

    public static DataSet executeQueryS(string sql, string connStr)
    {
        string strError = "";
        DataSet ds = new DataSet();
        SqlConnection conn = new SqlConnection(connStr);
        if (conn.State == ConnectionState.Closed)
        {
            conn.Open();
        }
        try
        {
            SqlDataAdapter mySqlDataAdapter = new SqlDataAdapter();
            SqlCommand selectCommand = new SqlCommand(sql, conn);
            selectCommand.CommandType = CommandType.Text;
            mySqlDataAdapter.SelectCommand = selectCommand;
            mySqlDataAdapter.Fill(ds);
            return ds;
        }
        catch (Exception ex)
        {
            strError = "数据检索失败：" + ex.Message;
            return null; 
        }
        finally
        {
            if (conn.State != ConnectionState.Closed)
            {
                conn.Close();                
            } 
        }
    }
    

   

    public static bool existInq(string table,string column,string columnText,string connStr)
    {
        string strSql = "select * from "+table+"  where "+column+"='"+columnText+"'";       
        DataTable dt = Utils.executeQueryT(strSql,connStr);
        if (dt.Rows.Count > 0)
        {
            return true;
        }
        else
        {
            return false;
        }
        dt.Clear();
    }

    public static bool existInq(string table, string column1, string columnText1, string column2, string columnText2, string connStr)
    {
        string strSql = "select * from " + table + "  where " + column1 + "='" + columnText1 + "' and " + column2 + "='" + columnText2 + "' ";
        DataTable dt = Utils.executeQueryT(strSql, connStr);
        if (dt.Rows.Count > 0)
        {
            return true;
        }
        else
        {
            return false;
        }
        dt.Clear();
    }

    public static bool existInq(string table, string column1, string columnText1, string column2, string columnText2, string column3, string columnText3, string connStr)
    {
        string strSql = "select * from " + table + "  where " + column1 + "='" + columnText1 + "' and " + column2 + "='" + columnText2 + "' and " + column3 + "='" + columnText3 + "'  ";
        DataTable dt = Utils.executeQueryT(strSql, connStr);
        if (dt.Rows.Count > 0)
        {
            return true;
        }
        else
        {
            return false;
        }
        dt.Clear();
    }

    public static bool existInq(string table, string column1, string columnText1, string column2, string columnText2, string column3, string columnText3, string column4, string columnText4, string connStr)
    {
        string strSql = "select * from " + table + "  where " + column1 + "='" + columnText1 + "' and " + column2 + "='" + columnText2 + "' and " + column3 + "='" + columnText3 + "' and " + column4 + "='" + columnText4 + "'   ";
        DataTable dt = Utils.executeQueryT(strSql, connStr);
        if (dt.Rows.Count > 0)
        {
            return true;
        }
        else
        {
            return false;
        }
        dt.Clear();
    }

    public static bool existInq(string table, string column1, string columnText1, string column2, string columnText2, string column3, string columnText3, string column4, string columnText4, string column5, string columnText5, string connStr)
    {
        string strSql = "select * from " + table + "  where " + column1 + "='" + columnText1 + "' and " + column2 + "='" + columnText2 + "' and " + column3 + "='" + columnText3 + "' and " + column4 + "='" + columnText4 + "'  and " + column5 + "='" + columnText5 + "'   ";
        DataTable dt = Utils.executeQueryT(strSql, connStr);
        if (dt.Rows.Count > 0)
        {
            return true;
        }
        else
        {
            return false;
        }
        dt.Clear();
    }

    public static bool existInq(string table, string column1, string columnText1, string column2, string columnText2, string column3, string columnText3, string column4, string columnText4, string column5, string columnText5, string column6, string columnText6, string connStr)
    {
        string strSql = "select * from " + table + "  where " + column1 + "='" + columnText1 + "' and " + column2 + "='" + columnText2 + "' and " + column3 + "='" + columnText3 + "' and " + column4 + "='" + columnText4 + "' and " + column5 + "='" + columnText5 + "' and " + column6 + "='" + columnText6 + "'   ";
        DataTable dt = Utils.executeQueryT(strSql, connStr);
        if (dt.Rows.Count > 0)
        {
            return true;
        }
        else
        {
            return false;
        }
        dt.Clear();
    }

    public static bool existInq(string table, string column1, string columnText1, string column2, string columnText2, string column3, string columnText3, string column4, string columnText4, string column5, string columnText5, string column6, string columnText6, string column7, string columnText7, string connStr)
    {
        string strSql = "select * from " + table + "  where " + column1 + "='" + columnText1 + "' and " + column2 + "='" + columnText2 + "' and " + column3 + "='" + columnText3 + "' and " + column4 + "='" + columnText4 + "' and " + column5 + "='" + columnText5 + "' and " + column6 + "='" + columnText6 + "'  and " + column7 + "='" + columnText7 + "'  ";
        DataTable dt = Utils.executeQueryT(strSql, connStr);
        if (dt.Rows.Count > 0)
        {
            return true;
        }
        else
        {
            return false;
        }
        dt.Clear();
    }

    public static bool existInq(string table, string column1, string columnText1, string column2, string columnText2, string column3, string columnText3, string column4, string columnText4, string column5, string columnText5, string column6, string columnText6, string column7, string columnText7, string column8, string columnText8, string connStr)
    {
        string strSql = "select * from " + table + "  where " + column1 + "='" + columnText1 + "' and " + column2 + "='" + columnText2 + "' and " + column3 + "='" + columnText3 + "' and " + column4 + "='" + columnText4 + "' and " + column5 + "='" + columnText5 + "' and " + column6 + "='" + columnText6 + "'  and " + column7 + "='" + columnText7 + "' and " + column8 + "='" + columnText8 + "'  ";
        DataTable dt = Utils.executeQueryT(strSql, connStr);
        if (dt.Rows.Count > 0)
        {
            return true;
        }
        else
        {
            return false;
        }
        dt.Clear();
    }



    public static void sendMail(string mailFrom, string mailTo, string subject, string message,string _Emp,string _cust,string _pay)
    {
        try
        {
            SmtpClient client = new SmtpClient("mail.sinno-tech.com");   //设置邮件协议
            client.UseDefaultCredentials = false;//这一句得写前面
            client.DeliveryMethod = SmtpDeliveryMethod.Network; //通过网络发送到Smtp服务器
            client.Credentials = new NetworkCredential("Justin.wang", "5t6y&U*I"); //通过用户名和密码 认证
            MailMessage mmsg = new MailMessage(new MailAddress("Finance@sinno-tech.com"), new MailAddress("Justin.wang@sinno-tech.com")); //发件人和收件人的邮箱地址
            mmsg.Subject = _cust+ "Overdue Payment Reminder";      //邮件主题
            mmsg.SubjectEncoding = Encoding.UTF8;   //主题编码
            mmsg.Body = "<p class=MsoNormal><span lang=EN-US>Dear "+_Emp+",<o:p></o:p></span></p><p class=MsoNormal><span lang=EN-US>&nbsp;&nbsp; </span><span style='font-family:宋体'>截至</span><span lang=EN-US>"+DateTime.Now.ToString("yyyy-MM-dd")+"</span><span style='font-family:宋体'>，"+_cust+"对我司尚存在未支付货款</span><span lang=EN-US>"+_pay+ "</span><span style='font-family:宋体'>，已超出约定付款期限（明细详见催款函），但我公司至今未收到该款项，故发函提醒，请予以确认。</span><span lang=EN-US><o:p></o:p></span></p><p class=MsoNormal><span lang=EN-US>&nbsp;&nbsp;</span><span style='font-family:宋体;color:#4472C4'>此邮件为系统自动发送请勿回复，谢谢！</span><p class=MsoNormal><o:p></o:p><o:p></o:p><span lang=EN-US>Finance</span></p>";         //邮件正文
            mmsg.BodyEncoding = Encoding.UTF8;      //正文编码
            mmsg.IsBodyHtml = true;    //设置为HTML格式          
            mmsg.Priority = MailPriority.High;   //优先级
            if(message.Trim().Length>0)
                mmsg.Attachments.Add(new Attachment(message));//增加附件          

            client.Send(mmsg); 
        }
        catch (Exception e)
        {
            Writelogstream(DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss") + " "+e.ToString());
        }
        finally
        {
            string _msg = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss")+" 已成功发送邮件至："+mailTo;
            Writelogstream(_msg);
        }
    }
    public static void Writelogstream(string _str)
    {
        System.IO.StreamWriter _sw = new System.IO.StreamWriter(System.Windows.Forms.Application.StartupPath + @"\Reminder.log", true);
        _sw.WriteLine(_str);
        _sw.Close();
    }
  

    
    public static string getPwdMd5(string str)
    {
        string pwd = "";
        System.Security.Cryptography.MD5 md5 = System.Security.Cryptography.MD5.Create();//实例化一个md5对像
        // 加密后是一个字节类型的数组，这里要注意编码UTF8/Unicode等的选择　
        byte[] s = md5.ComputeHash(System.Text.Encoding.UTF8.GetBytes(str));
        // 通过使用循环，将字节类型的数组转换为字符串，此字符串是常规字符格式化所得
        for (int i = 0; i < s.Length; i++)
        {
            // 将得到的字符串使用十六进制类型格式。格式后的字符是小写的字母，如果使用大写（X）则格式后的字符是大写字符 
            pwd = pwd + s[i].ToString("X");
        }
        return pwd;
    }

    public static bool CheckDate(string dayDate)
    {
        if (dayDate != "")
        {
            DateTime startDay;
            DateTime endDay;
            if (DateTime.TryParse(dayDate, out startDay))
            {
                return true;
            }
            else
            {
                return false;
            }
        }
        else
        {
            return true;
        }
    }

    public static bool CheckDate(string dateFrom, string DateTo)
    {
        if (dateFrom != "" && DateTo != "")
        {
            DateTime startDay;
            DateTime endDay;
            if (DateTime.TryParse(dateFrom, out startDay) && DateTime.TryParse(DateTo, out endDay))
            {
                if (startDay <= endDay)
                {
                    return true;
                }
                else
                {
                    return false;
                }
            }
            else
            {
                return false;
            }
        }
        else
        {
            if (dateFrom == "" && DateTo == "")
            {
                return true;
            }
            else
            {
                return false;
            }
        }
    }

    public static bool sessoinExist(object userID)
    {
        if (userID != null && userID != "")
        {
            return true;
        }
        else
        {
            return false;
        }
    }

    public static bool sessoinExist(object userID, string permissionID)
    {
        if (userID != null && userID != "" )
        {
            //if (Utils.setControl(userID.ToString().Trim(), permissionID))
            //{
            //    return true;
            //}
            //else
            //{
            //    return false;
            //}
            return true;
        }
        else
        {
            return false;
        }
    }

    public static string sessoinStr(object userID)
    {
        string str = "";
        if (userID != null && userID != "")
        {
            str = Utils.errorStr;
        }
        else
        {
            str = Utils.loginStr;
        }
        return str;
    }

    public static string sessoinStr(object userID, string permissionID)
    {
        string str = "";
        if (userID != null && userID != "")
        {
            //if (Utils.setControl(userID.ToString().Trim(), permissionID))
            //{
            //    str = "";
            //}
            //else
            //{
            //    str = "<script language='javascript' type='text/javascript'>window.open('http://222.92.41.188/Compex/ErrorPage.aspx','_self');</script>";
            //}
            str = Utils.loginStr;
        }
        else
        {
            str = Utils.loginStr;
        }
        return str;
    }

    public static string getTsn(string table, string column, string connStr)
    {
        string strSql = "select isnull(max(" + column + "),0)+1 as Tsn from " + table + " ";
        DataTable dt = Utils.executeQueryT(strSql, connStr);
        if (dt.Rows.Count > 0)
        {
            return dt.Rows[0][0].ToString().Trim();
        }
        else
        {
            return "0";
        }
        dt.Clear();
    }

    public static string getTsn(string table, string column1,string columnText1,string column2, string connStr)
    {
        string strSql = "select isnull(max(" + column2 + "),0)+1 as Tsn from " + table + " where " + column1 + "='" + columnText1 + "'";
        DataTable dt = Utils.executeQueryT(strSql, connStr);
        if (dt.Rows.Count > 0)
        {
            return dt.Rows[0][0].ToString().Trim();
        }
        else
        {
            return "0";
        }
        dt.Clear();
    }

    public static string getTsn(string table, string column1, string columnText1, string column2, string columnText2, string column3, string connStr)
    {
        string strSql = "select isnull(max(" + column3 + "),0)+1 as Tsn from " + table + " where " + column1 + "='" + columnText1 + "' and " + column2 + "='" + columnText2 + "' ";
        DataTable dt = Utils.executeQueryT(strSql, connStr);
        if (dt.Rows.Count > 0)
        {
            return dt.Rows[0][0].ToString().Trim();
        }
        else
        {
            return "0";
        }
        dt.Clear();
    }

    public static string getTsn(string table, string column1, string columnText1, string column2, string columnText2, string column3, string columnText3, string column4, string connStr)
    {
        string strSql = "select isnull(max(" + column4 + "),0)+1 as Tsn from " + table + " where " + column1 + "='" + columnText1 + "' and " + column2 + "='" + columnText2 + "'  and " + column3 + "='" + columnText3 + "' ";
        DataTable dt = Utils.executeQueryT(strSql, connStr);
        if (dt.Rows.Count > 0)
        {
            return dt.Rows[0][0].ToString().Trim();
        }
        else
        {
            return "0";
        }
        dt.Clear();
    }

    public static string getColumnValue(string table, string column1, string columnText1, string column,string connStr)
    {
        string strSql = "select "+column+" from " + table + "  where " + column1 + "='" + columnText1 + "' ";
        DataTable dt = Utils.executeQueryT(strSql, connStr);
        if (dt.Rows.Count > 0)
        {
            return dt.Rows[0][0].ToString().Trim();
        }
        else
        {
            return "";
        }
        dt.Clear();
    }

    public static string getColumnValue(string table, string column1, string columnText1, string column2, string columnText2, string column, string connStr)
    {
        string strSql = "select " + column + " from " + table + "  where " + column1 + "='" + columnText1 + "' and " + column2 + "='" + columnText2 + "' ";
        DataTable dt = Utils.executeQueryT(strSql, connStr);
        if (dt.Rows.Count > 0)
        {
            return dt.Rows[0][0].ToString().Trim();
        }
        else
        {
            return "";
        }
        dt.Clear();
    }

    public static string getColumnValue(string table, string column1, string columnText1, string column2, string columnText2, string column3, string columnText3, string column, string connStr)
    {
        string strSql = "select " + column + " from " + table + "  where " + column1 + "='" + columnText1 + "' and " + column2 + "='" + columnText2 + "' and " + column3+ "='" + columnText3 + "' ";
        DataTable dt = Utils.executeQueryT(strSql, connStr);
        if (dt.Rows.Count > 0)
        {
            return dt.Rows[0][0].ToString().Trim();
        }
        else
        {
            return "";
        }
        dt.Clear();
    }

    public static string getColumnValue(string table, string column1, string columnText1, string column2, string columnText2, string column3, string columnText3, string column4, string columnText4, string column, string connStr)
    {
        string strSql = "select " + column + " from " + table + "  where " + column1 + "='" + columnText1 + "' and " + column2 + "='" + columnText2 + "' and " + column3 + "='" + columnText3 + "'  and " + column4 + "='" + columnText4 + "' ";
        DataTable dt = Utils.executeQueryT(strSql, connStr);
        if (dt.Rows.Count > 0)
        {
            return dt.Rows[0][0].ToString().Trim();
        }
        else
        {
            return "";
        }
        dt.Clear();
    }
}
