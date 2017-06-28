using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using log4net;

using Quartz;
using Quartz.Impl;
using Quartz.Impl.Triggers;

namespace PowerMonitor
{
    public partial class Form1 : Form
    {
        private static ILog _logger = LogManager.GetLogger(typeof(Form1));
        public Form1()
        {
            InitializeComponent();
            string autoNotice = ConfigHelper.GetAppConfig("autoNotice");
            _logger.Info("是否自动执行标识:" + autoNotice);
            //是否启动定时任务
            if ("1".Equals(autoNotice))
            {
                InitScheduleTask();
            }
			_logger.Info("监控程序启动完成。");            
        }
        
        /// <summary>
        /// 初始化定时任务
        /// </summary>
        private static void InitScheduleTask()
        {
            try
            {
                _logger.Info("初始化定时任务开始。");
                int hour = 12;
                int minute = 00;
                string taskTime = ConfigHelper.GetAppConfig("taskTime");
                if("".Equals(taskTime))
                {
                	_logger.Warn("未配置每天自动执行时间。");
                }else
                {
                	string[] times = taskTime.Split(':');
                	if(null != times && times.Length > 0)
                	{
                		hour = int.Parse(times[0]);
                		minute = int.Parse(times[1]);
                	}
                }
                //1.首先创建一个作业调度池
                ISchedulerFactory schedf = new StdSchedulerFactory();
                IScheduler sched = schedf.GetScheduler();
                //2.创建出来一个具体的作业
                IJobDetail job = JobBuilder.Create<NotifyJob>().Build();
                //3.创建并配置一个触发器
                ISimpleTrigger trigger = (ISimpleTrigger)TriggerBuilder.Create().WithSimpleSchedule(x => x.WithIntervalInSeconds(3).WithRepeatCount(int.MaxValue)).Build();

                ITrigger trigger2 = TriggerBuilder.Create()
                  .WithIdentity("myTrigger", "group2")
                  .ForJob(job)
                  .WithSchedule(CronScheduleBuilder.DailyAtHourAndMinute(hour, minute)) // 每天9:30执行一次
                    //.ModifiedByCalendar("myHolidays") // but not on holidays 设置那一天不知道
                  .Build();

                //4.加入作业调度池中
                sched.ScheduleJob(job, trigger2);
                //5.开始运行
                sched.Start();
                //QuartzManager.AddJob<TaskJob>("每隔5秒", "*/5 * * * * ?");//每隔5秒执行一次这个方法
                _logger.Info("初始化定时任务成功，每天执行时间：" + taskTime);
            }
            catch (Exception ex)
            {
                //Console.WriteLine(ex.Message);
                _logger.Error("初始化定时任务时发生异常:" + ex.Message + ex.StackTrace);
            }
        }

        /// <summary>
        /// 页面加载
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void Form1_Load(object sender, EventArgs e)
        {
            //初始化异常类别控件，添加项，Web控件DropDownList有对应的ListItem
            ListItem listItem0 = new ListItem("all", "全部");
            comboBox1.Items.Add(listItem0);
            string exceptionCode = ConfigHelper.GetValue("exceptionCode");
            _logger.Info("系统配置的异常类别为：" + exceptionCode);
            string[] exceptionCodes = exceptionCode.Split(',');
            if(null != exceptionCodes && exceptionCodes.Length > 0)
            {
            	foreach (string codeStr in exceptionCodes)
            	{
            		string[] code = codeStr.Split(':');
                    string id = code[0];
                    string name = code[1];
                    ListItem listItem = new ListItem(id, name);
		            comboBox1.Items.Add(listItem);
            	}
            }else
            {
            	ListItem listItem1 = new ListItem("0301", "电能表开盖");
	            ListItem listItem2 = new ListItem("0205", "电流失流");
	            comboBox1.Items.Add(listItem1);
	            comboBox1.Items.Add(listItem2);
            }
            
            //设置默认选择项，DropDownList会默认选择第一项。
            comboBox1.SelectedIndex = 0;//设置第一项为默认选择项。
            //comboBox1.SelectedItem = listItem1;//设置指定的项为默认选择项
            //初始化时间控件
            DateTime now = DateTime.Now;
			DateTime firstDay = new DateTime(now.Year, now.Month, 1);
			DateTime lastDay = now;
			if(now.Day > 1)
			{
				lastDay = now.AddDays(-1);
			}
 			dateTimePicker1.Value = firstDay;
 			dateTimePicker2.Value = lastDay;
        }

        /// <summary>
        /// 执行调用方法
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void button1_Click(object sender, EventArgs e)
        {
        	//判断选择日期
            if(DateTime.Compare(dateTimePicker1.Value, dateTimePicker2.Value) > 0)
            {
            	MessageBox.Show("开始日期大于结束日期，请重新选择！", "提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
            	return;
            }
            //获取选择的异常类别
            ListItem selectedItem = (ListItem)this.comboBox1.SelectedItem;
            string exceptionCodeArr = "";
            if ("all".Equals(selectedItem.ID))
            {
            	exceptionCodeArr = JobHelper.getExceptionCode();
            }
            else
            {
                exceptionCodeArr = selectedItem.ID;
                exceptionCodeArr = "(" + exceptionCodeArr + ")";
            }
            string beginTime = dateTimePicker1.Value.ToString("yyyy-MM-dd");
            string endTime = dateTimePicker2.Value.ToString("yyyy-MM-dd");
            JobHelper.ExecuteTask(exceptionCodeArr, beginTime, endTime);
            MessageBox.Show("任务执行结束！", "提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }

        private void comboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
            
        }

        private void label1_Click(object sender, EventArgs e)
        {

        }

        private void monthCalendar1_DateChanged(object sender, DateRangeEventArgs e)
        {

        }

        private void dateTimePicker1_ValueChanged(object sender, EventArgs e)
        {

        }
        
        private void dateTimePicker2_ValueChanged(object sender, EventArgs e)
        {

        }

        private void label3_Click(object sender, EventArgs e)
        {

        }
        
        private void Form1_SizeChanged(object sender, EventArgs e)
        {
            //判断是否选择的是最小化按钮 
            if (WindowState == FormWindowState.Minimized)
            {
                //隐藏任务栏区图标 
                this.ShowInTaskbar = false;
                //图标显示在托盘区 
                notifyIcon1.Visible = true;
            }
        }

        private void notifyIcon1_MouseDoubleClick(object sender, MouseEventArgs e)
        {
            if (WindowState == FormWindowState.Minimized)
            {
                //还原窗体显示    
                WindowState = FormWindowState.Normal;
                //激活窗体并给予它焦点
                this.Activate();
                //任务栏区显示图标
                this.ShowInTaskbar = true;
                //托盘区图标隐藏
                notifyIcon1.Visible = false;
            }
        }

        /// <summary>
        /// 确认是否退出
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void Form1_FormClosing(object sender, FormClosingEventArgs e)
        {
            
            if (MessageBox.Show("是否确认退出程序？", "退出", MessageBoxButtons.OKCancel, MessageBoxIcon.Question) == DialogResult.OK)
            {
            	_logger.Info("监控程序退出。");
                // 关闭所有的线程
                this.Dispose();
                this.Close();
                System.Environment.Exit(0); 
            }
            else
            {
                e.Cancel = true;
                //任务栏区显示图标
                this.ShowInTaskbar = true;
                //托盘区图标隐藏
                notifyIcon1.Visible = false;
            }
        }
    }

    public class NotifyJob : IJob
    {
        private static ILog _logger = LogManager.GetLogger(typeof(NotifyJob));
        /// <summary>
        /// 作业调度每次定时执行方法
        /// </summary>
        /// <param name="context"></param>
        public void Execute(IJobExecutionContext context)
        {
            _logger.Info("定时任务执行开始:" + DateTime.Now.ToString(""));
            JobHelper.RunJob();
            _logger.Info("定时任务执行结束:" + DateTime.Now.ToString(""));
        }
    }
}
