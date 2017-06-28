using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Data;
using System.IO;
using Oracle.ManagedDataAccess.Client;
using log4net;

namespace PowerMonitor
{
    public static class JobHelper
    {
        private static ILog _logger = LogManager.GetLogger(typeof(JobHelper));
        
        public static void SendEmail(string emailBody)
        {
            Email email = new Email();
            email.mailFrom = ConfigHelper.GetValue("emailSender");
            email.mailPwd = ConfigHelper.GetValue("senderPassword");
            if(!Validator.IsEmail(email.mailFrom))
            {
            	_logger.Warn("发送人邮箱地址未配置，或配置错误:" + email.mailFrom);
            	return;
            }
            email.mailSubject = ConfigHelper.GetValue("emailSubject");//"邮件主题";
            email.mailBody = emailBody;
            email.isbodyHtml = true;    //是否是HTML
            email.host = ConfigHelper.GetValue("smtpHost");//"smtp.126.com";
            email.mailToArray = ConfigHelper.GetValue("emailReceiver").Split(',');//接收者邮件集合
            if (email.Send)
            {
				_logger.Info("发送邮件成功。");
            }
            else
            {
				_logger.Info("发送邮件失败。");
            }
        }
        
        public static void RunJob()
        {
            string exceptionCode = JobHelper.getExceptionCode();
            //计算查询时间范围
            DateTime now = DateTime.Now;
            DateTime lastDay = now.AddDays(-1);
            string beginTime = lastDay.ToString("yyyy-MM-dd");
            string endTime = now.ToString("yyyy-MM-dd");
            ExecuteTask(exceptionCode, beginTime, endTime);
        }

        public static void ExecuteTask(string exceptionCode, string beginTime, string endTime)
        {
        	//获取配置中的供电单位编号
        	string orgNo = ConfigHelper.GetValue("orgNo");
        	if("".Equals(orgNo))
        	{
        		_logger.Warn("未配置供电单位编号。");
        		return;
        	}else{
        		orgNo = "'" + orgNo + "%'";
        	}
        	beginTime = "'" + beginTime + "'";
        	endTime = "'" + endTime + "'";
        	_logger.Info("任务参数，异常类别：" + exceptionCode + "起始时间：" + beginTime  + "截止时间：" + endTime + "供电单位编号：" + orgNo);
        	DataTable powerTb = getPowerData(exceptionCode, beginTime, endTime, orgNo);
            //List<PowerEntity> powerList = OracleHelper.TableToEntity<PowerEntity>(powerTb);
            //从配置中获取SQL语句
            string sql = ConfigHelper.GetValue("sqlStatement");
            if("".Equals(sql)){
            	_logger.Warn("未配置查询数据SQL语句。");
        		return;
            }
            sql = sql.Replace(":exceptionCode", exceptionCode);
            sql = sql.Replace(":beginTime", beginTime);
            sql = sql.Replace(":endTime", endTime);
            sql = sql.Replace(":orgNo", orgNo);
            /*
            OracleParameter[] param = new OracleParameter[4] { new OracleParameter(":exceptionCode", OracleDbType.Varchar2), new OracleParameter(":beginTime", OracleDbType.Varchar2), new OracleParameter(":endTime", OracleDbType.Varchar2), new OracleParameter(":orgNo", OracleDbType.Varchar2) };
            param[0].Value = exceptionCode;
            param[1].Value = beginTime;
            param[2].Value = endTime;
            param[3].Value = orgNo;
            List<PowerEntity> powerList = OracleHelper.ExecuteList<PowerEntity>(sql, System.Data.CommandType.Text, param);
            */
            
            if(powerTb == null || powerTb.Rows == null || powerTb.Rows.Count < 1)
            {
            	_logger.Info("执行任务查询无异常明细记录。");
                if (File.Exists("testdata.xls"))
                {
                    powerTb = ExcelHelper.InputFromExcel("testdata.xls", "SQL Results");
                }
            	//return;
            }
            string excelPath = getExcelPath();
            _logger.Info("导出Excel文件的路径:" + excelPath);
            string fileName = excelPath + "异常明细" + DateTime.Now.ToString("yyyy-MM-dd-HH-mm-ss") + ".xls";
            ExcelHelper.DatatableToExcelByNPOI(powerTb, fileName);
            DataTable anlysisTable = AnalysisTable(powerTb);
            if(anlysisTable != null && anlysisTable.Rows != null && anlysisTable.Rows.Count > 0)
            {
            	string emailBody = ExcelHelper.GetHtmlString("用电异常提醒", anlysisTable);
                SendEmail(emailBody);
            }else
            {
            	_logger.Info("没有信息需要发送。");
            }
        }

        public  static string getExceptionCode()
        {
        	string exceptionCodeArr = "";
        	string exceptionCode = ConfigHelper.GetValue("exceptionCode");
            string[] exceptionCodes = exceptionCode.Split(',');
            if (null != exceptionCodes && exceptionCodes.Length > 0)
            {
            	_logger.Info("系统配置的异常类别为：" + exceptionCode);
                foreach (string codeStr in exceptionCodes)
                {
                    string[] code = codeStr.Split(':');
                    string id = code[0];
                    if (codeStr.Equals(exceptionCodes[exceptionCodes.Length - 1]))
                    {
                        exceptionCodeArr += "'" + id + "'";
                    }
                    else
                    {
                        exceptionCodeArr += "'" + id + "'" + ",";
                    }
                } 
            }else
            {
            	_logger.Warn("未配置异常用电类型。");
                exceptionCodeArr = "0301" + "," + "0205";
            }
            //exceptionCodeArr = "(" + exceptionCodeArr + ")";
            return exceptionCodeArr;
        }
        
        public static string getExcelPath()
        {
        	string excelPath = ConfigHelper.GetValue("excelPath");
        	if("".Equals(excelPath))
        	{
        		_logger.Warn("未配置导出Excel文件的路径。");
        		excelPath = "temp";
        	}
        	try
            {
        		if (Directory.Exists(excelPath) == false)//如果不存在就创建file文件夹
		        {
		             Directory.CreateDirectory(excelPath);
		        }
            }
            catch (Exception ex)
            {
            	_logger.Error("创建导出Excel文件的路径时报错：" + ex.Message + ex.StackTrace);
            	return "";
            }
            return excelPath;
        }
        
        public static string BuilderTable(DataTable powerTb)
        {
        	StringBuilder bodyBuilder = new StringBuilder();
            bodyBuilder.Append("<table class=\"tableheadstyle\">");
            foreach (DataRow dr in powerTb.Rows)
            {
                bodyBuilder.Append("<tr>");
                bodyBuilder.AppendFormat("<td align=\"right\" style=\"background-color:#f7fbff;\"><div style=\"width:80px;\">{0}</div></td>", dr["USER_ID"].ToString());
                bodyBuilder.AppendFormat("<td align=\"right\" style=\"background-color:#f7fbff;\"><div style=\"width:80px;\">{0}</div></td>", dr["USER_NAME"].ToString());
                bodyBuilder.Append("</tr>");
            }
            bodyBuilder.Append("</table>");
            string resultStr = bodyBuilder.ToString();
            return resultStr;
        }

        public static DataTable getPowerData(string exceptionCode, string beginTime, string endTime, string orgNo)
        {
        	string constr = ConfigHelper.GetValue("connstr");
        	if("".Equals(constr)){
            	_logger.Warn("未配置ORACLE数据库连接信息。");
        		return null;
            }
			string sql = ConfigHelper.GetValue("sqlStatement");
            if("".Equals(sql)){
            	_logger.Warn("未配置查询数据SQL语句。");
        		return null;
            }
			sql = sql.Replace(":exceptionCode", exceptionCode);
            sql = sql.Replace(":beginTime", beginTime);
            sql = sql.Replace(":endTime", endTime);
            sql = sql.Replace(":orgNo", orgNo);
			_logger.Info("任务执行SQL语句:" + sql);
			DataTable tb = new DataTable();
			using (OracleConnection con = new OracleConnection(constr))
			{
			   try  
               {  
                   con.Open();   
                   using (OracleCommand cmd = new OracleCommand())
                   {
	                   	cmd.CommandType = System.Data.CommandType.Text;
			            cmd.CommandText = sql;
			            /*cmd.Parameters.Add(new OracleParameter(":exceptionCode",exceptionCode));
			            cmd.Parameters.Add(new OracleParameter(":beginTime",beginTime));
			            cmd.Parameters.Add(new OracleParameter(":endTime",endTime));
			            cmd.Parameters.Add(new OracleParameter(":orgNo",orgNo + "%"));*/
			        	cmd.Connection = con;	
			        	using (OracleDataAdapter da = new OracleDataAdapter())
			        	{
			        		da.SelectCommand = cmd;
				        	DataSet ds = new DataSet();
				        	da.Fill(ds);
				        	tb = ds.Tables[0];
			        	}
                   }
               }  
               catch (OracleException ex)  
               {  
                   //throw new Exception(ex.Message); 
                   _logger.Error("查询数据时发生异常:" + ex.Message + Environment.NewLine + ex.StackTrace);
               }  
			}
			
        	/*OracleConnection con = OracleHelper.OpenConn();
        	OracleCommand cmd = new OracleCommand();
        	cmd.CommandType = System.Data.CommandType.Text;
            cmd.CommandText = sql;
            cmd.Parameters.Add(new OracleParameter(":exceptionCode",exceptionCode));
            cmd.Parameters.Add(new OracleParameter(":beginTime",beginTime));
            cmd.Parameters.Add(new OracleParameter(":endTime",endTime));
            cmd.Parameters.Add(new OracleParameter(":orgNo",orgNo + "%"));
        	cmd.Connection = con;
        	
        	OracleDataAdapter da = new OracleDataAdapter();
        	da.SelectCommand = cmd;
        	DataSet ds = new DataSet();
        	da.Fill(ds);
        	DataTable tb = ds.Tables[0];*/
            return tb;    	
        }
        
        public static DataTable AnalysisTable(DataTable dt)
        {
        	_logger.Info("统计分析异常信息开始。");
        	DataTable analysisTable = new DataTable("analysisTable"); 
        	analysisTable.Columns.Add("供电单位", System.Type.GetType("System.String"));
        	analysisTable.Columns.Add("户号", System.Type.GetType("System.String"));
        	analysisTable.Columns.Add("户名", System.Type.GetType("System.String"));
        	analysisTable.Columns.Add("用电地址", System.Type.GetType("System.String"));
        	analysisTable.Columns.Add("电压等级", System.Type.GetType("System.String"));
        	analysisTable.Columns.Add("电能表条码号", System.Type.GetType("System.String"));
        	analysisTable.Columns.Add("是否电能表开盖", System.Type.GetType("System.String"));
        	analysisTable.Columns.Add("电能表开盖时间", System.Type.GetType("System.String"));
        	analysisTable.Columns.Add("是否电流失流", System.Type.GetType("System.String"));
        	analysisTable.Columns.Add("电流失流时间", System.Type.GetType("System.String"));
        	
        	Dictionary<string, DataRow> myDictionary = new Dictionary<string, DataRow>(); 
            try
            {
                //建立内容行
                foreach (DataRow Rowitem in dt.Rows)
                {
                	string cons_no = Rowitem["户号"].ToString();
                	if(myDictionary.ContainsKey(cons_no)) 
				　　{ 
                		DataRow dr = myDictionary[cons_no];
                		if("0301".Equals(Rowitem["异常类型代码"].ToString()))
                		{
                			dr["是否电能表开盖"] = "是";
                			dr["电能表开盖时间"] = Rowitem["末次告警时间"].ToString();
                		}else if("0205".Equals(Rowitem["异常类型代码"].ToString()))
                		{
                			dr["是否电流失流"] = "是";
                			dr["电流失流时间"] = Rowitem["末次告警时间"].ToString();
                		}
				　　} 
                	else
                	{
                		DataRow dr = analysisTable.NewRow();
                		dr["供电单位"] = Rowitem["供电单位"].ToString();
                		dr["户号"] = Rowitem["户号"].ToString();
                		dr["户名"] = Rowitem["户名"].ToString();
                		dr["用电地址"] = Rowitem["用电地址"].ToString();
                		dr["电压等级"] = Rowitem["电压等级"].ToString();
                		dr["电能表条码号"] = Rowitem["电能表条码号"].ToString();
                		if("0301".Equals(Rowitem["异常类型代码"].ToString()))
                		{
                			dr["是否电能表开盖"] = "是";
                			dr["电能表开盖时间"] = Rowitem["末次告警时间"].ToString();
                			/*
                			if(null != (Rowitem["末次告警时间"].ToString()) && !"".Equals(Rowitem["末次告警时间"].ToString()))
                			{
                				dr["电能表开盖时间"] = Rowitem["末次告警时间"].ToString();
                			}else
                			{
                				dr["电能表开盖时间"] = Rowitem["首次告警时间"].ToString();
                			}
                			*/
                		}else if("0205".Equals(Rowitem["异常类型代码"].ToString()))
                		{
                			dr["是否电流失流"] = "是";
                			dr["电流失流时间"] = Rowitem["末次告警时间"].ToString();
                		}
                		myDictionary.Add(cons_no, dr);
                		analysisTable.Rows.Add(dr);
                	}
                }
            }
            catch (Exception ex)
            {
                _logger.Error("统计分析数据是发生异常：" + ex.Message + Environment.NewLine + ex.StackTrace);
            }
            _logger.Info("统计分析异常信息完成。");
            return analysisTable;
        }
        
        /*
        private static void send_email()
        {
            var emailAcount = "yuli16888@126.com"; //ConfigurationManager.AppSettings["EmailAcount"];
            var emailPassword = "test";//ConfigurationManager.AppSettings["EmailPassword"];
            var reciver = "test@163.com";
            var content = "hello";
            MailMessage message = new MailMessage();
            //设置发件人,发件人需要与设置的邮件发送服务器的邮箱一致
            MailAddress fromAddr = new MailAddress("yuli16888@126.com");
            message.From = fromAddr;
            //设置收件人,可添加多个,添加方法与下面的一样
            message.To.Add(reciver);
            //设置抄送人
            message.CC.Add("yuli16888@126.com");
            //设置邮件标题
            message.Subject = "Test";
            //设置邮件内容
            message.Body = content;
            //设置邮件发送服务器,服务器根据你使用的邮箱而不同,可以到相应的 邮箱管理后台查看,下面是QQ的
            //SmtpClient client = new SmtpClient("smtp.126.com", 25);
            SmtpClient client = new SmtpClient();
            client.Host = "smtp.126.com";
            //设置发送人的邮箱账号和密码
            client.Credentials = new NetworkCredential(emailAcount, emailPassword);
            //启用ssl,也就是安全发送
            client.EnableSsl = true;
            //发送邮件
            client.Send(message);
        }*/
    }
}
