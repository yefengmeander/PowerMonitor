using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Oracle.ManagedDataAccess.Client;
using System.Configuration;
using System.Data;
using System.Reflection;
namespace PowerMonitor
{
    public class OracleHelper
    {
        #region 变量  

        /// <summary>  
        /// 数据库连接对象  
        /// </summary>  
        private static OracleConnection _con = null;
        public static string constr = ConfigHelper.GetValue("connstr");//ConfigurationManager.ConnectionStrings["OracleStr"].ToString();
        #endregion


        #region 属性  
        
        public static OracleConnection OpenConn()
        {
            OracleConnection conn = new OracleConnection();
            //conn.ConnectionString = "Data Source=(DESCRIPTION=(ADDRESS=(PROTOCOL=TCP)(HOST=***.***.***.***)(PORT=1521))(CONNECT_DATA=(SERVICE_NAME=***)));Persist Security Info=True;User ID=***;Password=***;";
            conn.ConnectionString = ConfigHelper.GetValue("connstr");
            try
            {
                conn.Open();
            }
            catch (Exception e)
            {
                Console.WriteLine(e.Message);
            }
            return conn;
        }

        public static void CloseConn(OracleConnection conn)
        {
            if (conn == null) { return; }
            try
            {
                if (conn.State != ConnectionState.Closed)
                {
                    conn.Close();
                }
            }
            catch (Exception e)
            {
                Console.WriteLine(e.Message);
            }
            finally
            {
                conn.Dispose();
            }
        }

        /// <summary>  

        /// 获取或设置数据库连接对象  

        /// </summary>  

        public static OracleConnection Con
        {

            get
            {

                if (OracleHelper._con == null)
                {

                    OracleHelper._con = new OracleConnection();

                }

                if (OracleHelper._con.ConnectionString == "")
                {

                    OracleHelper._con.ConnectionString = OracleHelper.constr;

                }

                return OracleHelper._con;

            }
            set
            {
                OracleHelper._con = value;
            }
        }
        #endregion

        #region 方法  

        #region 执行返回一行一列的数据库操作  

        /// <summary>  

        /// 执行返回一行一列的数据库操作  

        /// </summary>  

        /// <param name="commandText">Oracle语句或存储过程名</param>  

        /// <param name="commandType">Oracle命令类型</param>  

        /// <param name="param">Oracle命令参数数组</param>  

        /// <returns>第一行第一列的记录</returns>  

        public static int ExecuteScalar(string commandText, CommandType commandType, params OracleParameter[] param)
        {

            int count = 0;

            using (OracleHelper.Con)
            {

                using (OracleCommand cmd = new OracleCommand(commandText, OracleHelper.Con))
                {

                    //try
                    //{

                        cmd.CommandType = commandType;
                        if (param != null)
                        {
                            cmd.Parameters.AddRange(param);
                        }


                        OracleHelper.Con.Open();

                        count = Convert.ToInt32(cmd.ExecuteScalar());

                    //}

                    //catch (Exception ex)
                    //{

                    //    count = 0;

                    //}

                }


            }

            return count;

        }

        #endregion


        #region 执行不查询的数据库操作  

        /// <summary>  

        /// 执行不查询的数据库操作  

        /// </summary>  

        /// <param name="commandText">Oracle语句或存储过程名</param>  

        /// <param name="commandType">Oracle命令类型</param>  

        /// <param name="param">Oracle命令参数数组</param>  

        /// <returns>受影响的行数</returns>  

        public static int ExecuteNonQuery(string commandText, CommandType commandType, params OracleParameter[] param)
        {

            int result = 0;

            using (OracleHelper.Con)
            {

                using (OracleCommand cmd = new OracleCommand(commandText, OracleHelper.Con))
                {

                    try
                    {

                        cmd.CommandType = commandType;
                        if (param != null)
                        {
                            cmd.Parameters.AddRange(param);
                        }



                        OracleHelper.Con.Open();

                        result = cmd.ExecuteNonQuery();

                    }

                    catch (Exception ex)
                    {

                        result = 0;

                    }

                }


            }

            return result;

        }

        #endregion


        #region 执行返回一条记录的泛型对象  

        /// <summary>  

        /// 执行返回一条记录的泛型对象  

        /// </summary>  

        /// <typeparam name="T">泛型类型</typeparam>  

        /// <param name="reader">只进只读对象</param>  

        /// <returns>泛型对象</returns>  

        private static T ExecuteDataReader<T>(IDataReader reader)
        {

            T obj = default(T);
            try
            {
                Type type = typeof(T);

                obj = (T)Activator.CreateInstance(type);//从当前程序集里面通过反射的方式创建指定类型的对象     

                //obj = (T)Assembly.Load(OracleHelper._assemblyName).CreateInstance(OracleHelper._assemblyName + "." + type.Name);//从另一个程序集里面通过反射的方式创建指定类型的对象   

                PropertyInfo[] propertyInfos = type.GetProperties();//获取指定类型里面的所有属性  

                foreach (PropertyInfo propertyInfo in propertyInfos)
                {
                    for (int i = 0; i < reader.FieldCount; i++)
                    {
                        string fieldName = reader.GetName(i);
                        if (fieldName.ToLower() == propertyInfo.Name.ToLower())
                        {
                            object val = reader[propertyInfo.Name];//读取表中某一条记录里面的某一列  
                            if (val != null && val != DBNull.Value)
                            {
                                if (val.GetType() == typeof(decimal) || val.GetType() == typeof(int))
                                {
                                    propertyInfo.SetValue(obj, Convert.ToInt32(val), null);
                                }
                                else if (val.GetType() == typeof(DateTime))
                                {
                                    propertyInfo.SetValue(obj, Convert.ToDateTime(val), null);
                                }
                                else if (val.GetType() == typeof(string))
                                {
                                    propertyInfo.SetValue(obj, Convert.ToString(val), null);
                                }
                            }
                            break;
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                throw;
            }
            return obj;
        }
        #endregion


        #region 执行返回一条记录的泛型对象  

        /// <summary>  

        /// 执行返回一条记录的泛型对象  

        /// </summary>  

        /// <typeparam name="T">泛型类型</typeparam>  

        /// <param name="commandText">Oracle语句或存储过程名</param>  

        /// <param name="commandType">Oracle命令类型</param>  

        /// <param name="param">Oracle命令参数数组</param>  

        /// <returns>实体对象</returns>  

        public static T ExecuteEntity<T>(string commandText, CommandType commandType, params OracleParameter[] param)
        {

            T obj = default(T);

            using (OracleHelper.Con)
            {

                using (OracleCommand cmd = new OracleCommand(commandText, OracleHelper.Con))
                {

                    cmd.CommandType = commandType;

                    cmd.Parameters.AddRange(param);

                    OracleHelper.Con.Open();

                    OracleDataReader reader = cmd.ExecuteReader(CommandBehavior.CloseConnection);

                    while (reader.Read())
                    {

                        obj = OracleHelper.ExecuteDataReader<T>(reader);

                    }

                }

            }

            return obj;

        }

        #endregion


        #region 执行返回多条记录的泛型集合对象  

        /// <summary>  

        /// 执行返回多条记录的泛型集合对象  

        /// </summary>  

        /// <typeparam name="T">泛型类型</typeparam>  

        /// <param name="commandText">Oracle语句或存储过程名</param>  

        /// <param name="commandType">Oracle命令类型</param>  

        /// <param name="param">Oracle命令参数数组</param>  

        /// <returns>泛型集合对象</returns>  

        public static List<T> ExecuteList<T>(string commandText, CommandType commandType, params OracleParameter[] param)
        {

            List<T> list = new List<T>();

            using (OracleHelper.Con)
            {

                using (OracleCommand cmd = new OracleCommand(commandText, OracleHelper.Con))
                {

                    try
                    {

                        cmd.CommandType = commandType;
                        if (param != null)
                        {
                            cmd.Parameters.AddRange(param);
                        }
                        OracleHelper.Con.Open();

                        OracleDataReader reader = cmd.ExecuteReader(CommandBehavior.CloseConnection);

                        while (reader.Read())
                        {


                            T obj = OracleHelper.ExecuteDataReader<T>(reader);

                            list.Add(obj);

                        }

                    }

                    catch (Exception ex)
                    {

                        list = null;
                    }
                }
            }
            return list;
        }
        #endregion

        #endregion
        
        //List<User> userList = TableToEntity<User>(YourDataTable);

        public static List<T> TableToEntity<T>(DataTable dt) where T : class,new() 
		{ 
		    Type type = typeof(T); 
		    List<T> list = new List<T>(); 
		
		    foreach (DataRow row in dt.Rows) 
		    { 
		        PropertyInfo[] pArray = type.GetProperties(); 
		        T entity = new T(); 
		        foreach (PropertyInfo p in pArray) 
		        { 
		            if (row[p.Name] is Int64) 
		            { 
		                p.SetValue(entity, Convert.ToInt32(row[p.Name]), null); 
		                continue; 
		            } 
		            p.SetValue(entity, row[p.Name], null); 
		        } 
		        list.Add(entity); 
		    } 
		    return list; 
		} 

    }
}
