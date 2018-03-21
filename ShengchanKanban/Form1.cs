using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Data.SqlClient;

namespace ShengchanKanban
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }
        
        private void Form1_Load(object sender, EventArgs e)
        {
            this.timer1.Interval = 30000;   //设置定时器30s
            this.timer1.Start();
                   
            //全屏显示无边框
            this.SetVisibleCore(false);
            this.FormBorderStyle = FormBorderStyle.None;
            this.WindowState = FormWindowState.Maximized;
            this.SetVisibleCore(true);

            DateTime dt = DateTime.Now;

            label1.Text = "SMT一号线日生产计划（The Daily Task ForSMT Line 1）/" + dt.ToShortDateString().ToString();
            this.label1.Location = new System.Drawing.Point(0, 0);
            this.label1.Height = 45;

            label2.Text = "SMT二号线日生产计划（The Daily Task For SMT Line 2）/" + dt.ToShortDateString().ToString();
            this.label2.Location = new System.Drawing.Point(0, 540);
            this.label2.Height = 45;


            string sql = string.Format(@"select * from (
	select distinct A.Fcode '计划任务单号', 
	case when A.Status=0 then '新建' when A.Status=1 then '保留' when A.Status=4 then '完成' when A.Status=5 then '指定结束' 
	when not exists(select null from SfcGreOperateList P inner join SfcGreWorkStationLog L on L.ID=WsLogId where P.SfcId=B.SfcId and PassType=1 and L.WorkStation='02.01.01') and a.status=2 then '未开工' 
	when exists(select null from SfcGreOperateList P inner join SfcGreWorkStationLog L on L.ID=WsLogId where P.SfcId=B.SfcId and PassType=2 and L.WorkStation='02.01.01') and a.status=2 then '完工' 
	when exists(select null from SfcGreOperateList P where P.SfcId=B.SfcId and Operate='FirstComp') and a.status=2 then '量产中' 
	when exists(select null from SfcGreOperateList P where P.SfcId=B.SfcId and Operate='Prepaired') and a.status=2 then '首件生产' 
	when exists(select null from SfcGreOperateList P inner join SfcGreWorkStationLog L on L.ID=WsLogId where P.SfcId=B.SfcId and PassType=1 and L.WorkStation='02.01.01' ) and a.status=2 then '准备中' 
	else '' end '状态', 
	A.Product '产品编号',M.Fname '产品名称',M.Fmodel '规格', 
	substring(convert(nvarchar(16),V.PreBeginTime,20),6,11) '计划开始时间', 
	substring(convert(nvarchar(16),V.PreEndTime,20),6,11) '计划结束时间', 
	Convert(float, A.FQty) '预计产量', 
	case when a.status in(4,5) or (exists(select null from SfcGreOperateList P inner join SfcGreWorkStationLog L on L.ID=WsLogId where P.SfcId=B.SfcId and PassType=2 and L.WorkStation='02.01.01') and a.status=2)  then Convert(float, A.FQty) else isnull(CQ,0) end '完成数量' 
	from SfcExpScheduleView V 
	left join SfcExpSchedule A on V.ScheID=A.ID 
	left join SfcGreNoList B on A.ID=B.ScheduleId 
	left join MdcDatMaterial M on M.Fcode=A.Product 
	left join (select L.ScheduleId,count(*) CQ from SmtGreCountScanData S 
	inner join SfcGreNoList L on S.SfcId=L.SfcId where Scanner='ZY0003' and Barcode<>'ER'  group by ScheduleId) C on C.ScheduleId=A.ID 
	where V.Area='02.01' AND  A.Status<>3 
	AND exists(select null from SfcGreOperateList P where P.SfcId=B.SfcId and Operate='FirstComp') and a.status=2
	and convert(nvarchar(10),V.PreBeginTime,120)!=convert(nvarchar(10),Getdate(),120)
) a
where a.状态 in ('量产中','首件生产')
union ALL
select distinct A.Fcode '计划任务单号', 
case when A.Status=0 then '新建' when A.Status=1 then '保留' when A.Status=4 then '完成' when A.Status=5 then '指定结束' 
when not exists(select null from SfcGreOperateList P inner join SfcGreWorkStationLog L on L.ID=WsLogId where P.SfcId=B.SfcId and PassType=1 and L.WorkStation='02.01.01') and a.status=2 then '未开工' 
when exists(select null from SfcGreOperateList P inner join SfcGreWorkStationLog L on L.ID=WsLogId where P.SfcId=B.SfcId and PassType=2 and L.WorkStation='02.01.01') and a.status=2 then '完工' 
when exists(select null from SfcGreOperateList P where P.SfcId=B.SfcId and Operate='FirstComp') and a.status=2 then '量产中' 
when exists(select null from SfcGreOperateList P where P.SfcId=B.SfcId and Operate='Prepaired') and a.status=2 then '首件生产' 
when exists(select null from SfcGreOperateList P inner join SfcGreWorkStationLog L on L.ID=WsLogId where P.SfcId=B.SfcId and PassType=1 and L.WorkStation='02.01.01' ) and a.status=2 then '准备中' 
else '' end '状态', 
A.Product '产品编号',M.Fname '产品名称',M.Fmodel '规格', 
substring(convert(nvarchar(16),V.PreBeginTime,20),6,11) '计划开始时间', 
substring(convert(nvarchar(16),V.PreEndTime,20),6,11) '计划结束时间', 
Convert(float, A.FQty) '预计产量', 
case when a.status in(4,5) or (exists(select null from SfcGreOperateList P inner join SfcGreWorkStationLog L on L.ID=WsLogId where P.SfcId=B.SfcId and PassType=2 and L.WorkStation='02.01.01') and a.status=2)  then Convert(float, A.FQty) else isnull(CQ,0) end '完成数量' 
from SfcExpScheduleView V 
left join SfcExpSchedule A on V.ScheID=A.ID 
left join SfcGreNoList B on A.ID=B.ScheduleId 
left join MdcDatMaterial M on M.Fcode=A.Product 
left join (select L.ScheduleId,count(*) CQ from SmtGreCountScanData S 
inner join SfcGreNoList L on S.SfcId=L.SfcId where Scanner='ZY0003' and Barcode<>'ER'  group by ScheduleId) C on C.ScheduleId=A.ID 
where V.Area='02.01' AND  A.Status<>3 
and convert(nvarchar(10),V.PreBeginTime,120)=convert(nvarchar(10),Getdate(),120)
order by '计划开始时间'
 ");

            

            string sql0 = string.Format(@"select * from (
	select distinct A.Fcode '计划任务单号', 
	case when A.Status=0 then '新建' when A.Status=1 then '保留' when A.Status=4 then '完成' when A.Status=5 then '指定结束' 
	when not exists(select null from SfcGreOperateList P inner join SfcGreWorkStationLog L on L.ID=WsLogId where P.SfcId=B.SfcId and PassType=1 and L.WorkStation='02.02.01') and a.status=2 then '未开工' 
	when exists(select null from SfcGreOperateList P inner join SfcGreWorkStationLog L on L.ID=WsLogId where P.SfcId=B.SfcId and PassType=2 and L.WorkStation='02.02.01') and a.status=2 then '完工' 
	when exists(select null from SfcGreOperateList P where P.SfcId=B.SfcId and Operate='FirstComp') and a.status=2 then '量产中' 
	when exists(select null from SfcGreOperateList P where P.SfcId=B.SfcId and Operate='Prepaired') and a.status=2 then '首件生产' 
	when exists(select null from SfcGreOperateList P inner join SfcGreWorkStationLog L on L.ID=WsLogId where P.SfcId=B.SfcId and PassType=1 and L.WorkStation='02.02.01' ) and a.status=2 then '准备中' 
	else '' end '状态', 
	A.Product '产品编号',M.Fname '产品名称',M.Fmodel '规格', 
	substring(convert(nvarchar(16),V.PreBeginTime,20),6,11) '计划开始时间', 
	substring(convert(nvarchar(16),V.PreEndTime,20),6,11) '计划结束时间', 
	Convert(float, A.FQty) '预计产量', 
	case when a.status in(4,5) or (exists(select null from SfcGreOperateList P inner join SfcGreWorkStationLog L on L.ID=WsLogId where P.SfcId=B.SfcId and PassType=2 and L.WorkStation='02.02.01') and a.status=2)  then Convert(float, A.FQty) else isnull(CQ,0) end '完成数量' 
	from SfcExpScheduleView V 
	left join SfcExpSchedule A on V.ScheID=A.ID 
	left join SfcGreNoList B on A.ID=B.ScheduleId 
	left join MdcDatMaterial M on M.Fcode=A.Product 
	left join (select L.ScheduleId,count(*) CQ from SmtGreCountScanData S 
	inner join SfcGreNoList L on S.SfcId=L.SfcId where Scanner='ZY0017' and Barcode<>'ER'  group by ScheduleId) C on C.ScheduleId=A.ID 
	where V.Area='02.02' AND  A.Status<>3 
	AND exists(select null from SfcGreOperateList P where P.SfcId=B.SfcId and Operate='FirstComp') and a.status=2
	and convert(nvarchar(10),V.PreBeginTime,120)!=convert(nvarchar(10),Getdate(),120)
) a
where a.状态 in ('量产中','首件生产')
union ALL
select distinct A.Fcode '计划任务单号', 
case when A.Status=0 then '新建' when A.Status=1 then '保留' when A.Status=4 then '完成' when A.Status=5 then '指定结束' 
when not exists(select null from SfcGreOperateList P inner join SfcGreWorkStationLog L on L.ID=WsLogId where P.SfcId=B.SfcId and PassType=1 and L.WorkStation='02.02.01') and a.status=2 then '未开工' 
when exists(select null from SfcGreOperateList P inner join SfcGreWorkStationLog L on L.ID=WsLogId where P.SfcId=B.SfcId and PassType=2 and L.WorkStation='02.02.01') and a.status=2 then '完工' 
when exists(select null from SfcGreOperateList P where P.SfcId=B.SfcId and Operate='FirstComp') and a.status=2 then '量产中' 
when exists(select null from SfcGreOperateList P where P.SfcId=B.SfcId and Operate='Prepaired') and a.status=2 then '首件生产' 
when exists(select null from SfcGreOperateList P inner join SfcGreWorkStationLog L on L.ID=WsLogId where P.SfcId=B.SfcId and PassType=1 and L.WorkStation='02.02.01' ) and a.status=2 then '准备中' 
else '' end '状态', 
A.Product '产品编号',M.Fname '产品名称',M.Fmodel '规格', 
substring(convert(nvarchar(16),V.PreBeginTime,20),6,11) '计划开始时间', 
substring(convert(nvarchar(16),V.PreEndTime,20),6,11) '计划结束时间', 
Convert(float, A.FQty) '预计产量', 
case when a.status in(4,5) or (exists(select null from SfcGreOperateList P inner join SfcGreWorkStationLog L on L.ID=WsLogId where P.SfcId=B.SfcId and PassType=2 and L.WorkStation='02.02.01') and a.status=2)  then Convert(float, A.FQty) else isnull(CQ,0) end '完成数量' 
from SfcExpScheduleView V 
left join SfcExpSchedule A on V.ScheID=A.ID 
left join SfcGreNoList B on A.ID=B.ScheduleId 
left join MdcDatMaterial M on M.Fcode=A.Product 
left join (select L.ScheduleId,count(*) CQ from SmtGreCountScanData S 
inner join SfcGreNoList L on S.SfcId=L.SfcId where Scanner='ZY0017' and Barcode<>'ER'  group by ScheduleId) C on C.ScheduleId=A.ID 
where V.Area='02.02' AND  A.Status<>3 
and convert(nvarchar(10),V.PreBeginTime,120)=convert(nvarchar(10),Getdate(),120)
order by '计划开始时间'
");
            //连接数据库，绑定表1数据
            SqlConnection conn = new SqlConnection("server=192.168.1.6;database=Exd2CZMestest;uid=sa;pwd=flybarrier");
            SqlDataAdapter sda = new SqlDataAdapter(sql, conn);
            DataSet ds = new DataSet();
            sda.Fill(ds);
            dataGridView1.DataSource = ds.Tables[0];

            //绑定表2数据
            SqlDataAdapter sda1 = new SqlDataAdapter(sql0, conn);
            DataSet ds1 = new DataSet();
            sda1.Fill(ds1);
            dataGridView2.DataSource = ds1.Tables[0];

            this.dataGridView1.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.AllCellsExceptHeaders;
            this.dataGridView1.RowHeadersVisible = false;//datagridview前面的空白部分去除
            this.dataGridView1.ScrollBars = ScrollBars.None;//滚动条去除
            dataGridView1.Location = new System.Drawing.Point(0, 50);
            dataGridView1.Height = 485;
            //取消选中第一行
            dataGridView1.Rows[0].Selected = false;
            //隐藏列
            dataGridView1.Columns[2].Visible = false;
            dataGridView1.Columns[3].Visible = false;
            dataGridView1.Columns[5].Visible = false;
            dataGridView1.Columns[6].Visible = false;
            //设置列宽
            dataGridView1.Columns[0].Width = 350;
            dataGridView1.Columns[1].Width = 250;
            dataGridView1.Columns[4].Width = 1020;
            dataGridView1.Columns[7].Width = 150;
            dataGridView1.Columns[8].Width = 150;

            //改变标题的高度;
            dataGridView1.ColumnHeadersHeightSizeMode = DataGridViewColumnHeadersHeightSizeMode.EnableResizing;
            dataGridView1.ColumnHeadersHeight = 40;
            //设置标题内容居中显示;  `
            dataGridView1.ColumnHeadersDefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
            dataGridView1.ColumnHeadersHeightSizeMode = DataGridViewColumnHeadersHeightSizeMode.DisableResizing;


           

            this.dataGridView2.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.AllCellsExceptHeaders;
            this.dataGridView2.RowHeadersVisible = false;//datagridview前面的空白部分去除
            this.dataGridView2.ScrollBars = ScrollBars.None;//滚动条去除
            dataGridView2.Location = new System.Drawing.Point(0, 590);
            this.dataGridView2.Height = 485;

            //取消选中第一行
            dataGridView2.Rows[0].Selected = false;
            //隐藏列
            dataGridView2.Columns[2].Visible = false;
            dataGridView2.Columns[3].Visible = false;
            dataGridView2.Columns[5].Visible = false;
            dataGridView2.Columns[6].Visible = false;
            //设置列宽
            this.dataGridView2.Columns[0].Width = 350;
            this.dataGridView2.Columns[1].Width = 250;
            this.dataGridView2.Columns[4].Width = 1020;
            this.dataGridView2.Columns[7].Width = 150;
            this.dataGridView2.Columns[8].Width = 150;
            
            //改变标题的高度;
            dataGridView2.ColumnHeadersHeightSizeMode = DataGridViewColumnHeadersHeightSizeMode.EnableResizing;
            dataGridView2.ColumnHeadersHeight = 40;
            //设置标题内容居中显示;  `
            dataGridView2.ColumnHeadersDefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
            dataGridView2.ColumnHeadersHeightSizeMode = DataGridViewColumnHeadersHeightSizeMode.DisableResizing;
        }
        
        private void timer1_Tick(object sender, EventArgs e)
        {     
            string sql = string.Format(@"select * from (
	select distinct A.Fcode '计划任务单号', 
	case when A.Status=0 then '新建' when A.Status=1 then '保留' when A.Status=4 then '完成' when A.Status=5 then '指定结束' 
	when not exists(select null from SfcGreOperateList P inner join SfcGreWorkStationLog L on L.ID=WsLogId where P.SfcId=B.SfcId and PassType=1 and L.WorkStation='02.01.01') and a.status=2 then '未开工' 
	when exists(select null from SfcGreOperateList P inner join SfcGreWorkStationLog L on L.ID=WsLogId where P.SfcId=B.SfcId and PassType=2 and L.WorkStation='02.01.01') and a.status=2 then '完工' 
	when exists(select null from SfcGreOperateList P where P.SfcId=B.SfcId and Operate='FirstComp') and a.status=2 then '量产中' 
	when exists(select null from SfcGreOperateList P where P.SfcId=B.SfcId and Operate='Prepaired') and a.status=2 then '首件生产' 
	when exists(select null from SfcGreOperateList P inner join SfcGreWorkStationLog L on L.ID=WsLogId where P.SfcId=B.SfcId and PassType=1 and L.WorkStation='02.01.01' ) and a.status=2 then '准备中' 
	else '' end '状态', 
	A.Product '产品编号',M.Fname '产品名称',M.Fmodel '规格', 
	substring(convert(nvarchar(16),V.PreBeginTime,20),6,11) '计划开始时间', 
	substring(convert(nvarchar(16),V.PreEndTime,20),6,11) '计划结束时间', 
	Convert(float, A.FQty) '预计产量', 
	case when a.status in(4,5) or (exists(select null from SfcGreOperateList P inner join SfcGreWorkStationLog L on L.ID=WsLogId where P.SfcId=B.SfcId and PassType=2 and L.WorkStation='02.01.01') and a.status=2)  then Convert(float, A.FQty) else isnull(CQ,0) end '完成数量' 
	from SfcExpScheduleView V 
	left join SfcExpSchedule A on V.ScheID=A.ID 
	left join SfcGreNoList B on A.ID=B.ScheduleId 
	left join MdcDatMaterial M on M.Fcode=A.Product 
	left join (select L.ScheduleId,count(*) CQ from SmtGreCountScanData S 
	inner join SfcGreNoList L on S.SfcId=L.SfcId where Scanner='ZY0003' and Barcode<>'ER'  group by ScheduleId) C on C.ScheduleId=A.ID 
	where V.Area='02.01' AND  A.Status<>3 
	AND exists(select null from SfcGreOperateList P where P.SfcId=B.SfcId and Operate='FirstComp') and a.status=2
	and convert(nvarchar(10),V.PreBeginTime,120)!=convert(nvarchar(10),Getdate(),120)
) a
where a.状态 in ('量产中','首件生产')
union ALL
select distinct A.Fcode '计划任务单号', 
case when A.Status=0 then '新建' when A.Status=1 then '保留' when A.Status=4 then '完成' when A.Status=5 then '指定结束' 
when not exists(select null from SfcGreOperateList P inner join SfcGreWorkStationLog L on L.ID=WsLogId where P.SfcId=B.SfcId and PassType=1 and L.WorkStation='02.01.01') and a.status=2 then '未开工' 
when exists(select null from SfcGreOperateList P inner join SfcGreWorkStationLog L on L.ID=WsLogId where P.SfcId=B.SfcId and PassType=2 and L.WorkStation='02.01.01') and a.status=2 then '完工' 
when exists(select null from SfcGreOperateList P where P.SfcId=B.SfcId and Operate='FirstComp') and a.status=2 then '量产中' 
when exists(select null from SfcGreOperateList P where P.SfcId=B.SfcId and Operate='Prepaired') and a.status=2 then '首件生产' 
when exists(select null from SfcGreOperateList P inner join SfcGreWorkStationLog L on L.ID=WsLogId where P.SfcId=B.SfcId and PassType=1 and L.WorkStation='02.01.01' ) and a.status=2 then '准备中' 
else '' end '状态', 
A.Product '产品编号',M.Fname '产品名称',M.Fmodel '规格', 
substring(convert(nvarchar(16),V.PreBeginTime,20),6,11) '计划开始时间', 
substring(convert(nvarchar(16),V.PreEndTime,20),6,11) '计划结束时间', 
Convert(float, A.FQty) '预计产量', 
case when a.status in(4,5) or (exists(select null from SfcGreOperateList P inner join SfcGreWorkStationLog L on L.ID=WsLogId where P.SfcId=B.SfcId and PassType=2 and L.WorkStation='02.01.01') and a.status=2)  then Convert(float, A.FQty) else isnull(CQ,0) end '完成数量' 
from SfcExpScheduleView V 
left join SfcExpSchedule A on V.ScheID=A.ID 
left join SfcGreNoList B on A.ID=B.ScheduleId 
left join MdcDatMaterial M on M.Fcode=A.Product 
left join (select L.ScheduleId,count(*) CQ from SmtGreCountScanData S 
inner join SfcGreNoList L on S.SfcId=L.SfcId where Scanner='ZY0003' and Barcode<>'ER'  group by ScheduleId) C on C.ScheduleId=A.ID 
where V.Area='02.01' AND  A.Status<>3 
and convert(nvarchar(10),V.PreBeginTime,120)=convert(nvarchar(10),Getdate(),120)
order by '计划开始时间'
 ");

            string sql0 = string.Format(@"select * from (
	select distinct A.Fcode '计划任务单号', 
	case when A.Status=0 then '新建' when A.Status=1 then '保留' when A.Status=4 then '完成' when A.Status=5 then '指定结束' 
	when not exists(select null from SfcGreOperateList P inner join SfcGreWorkStationLog L on L.ID=WsLogId where P.SfcId=B.SfcId and PassType=1 and L.WorkStation='02.02.01') and a.status=2 then '未开工' 
	when exists(select null from SfcGreOperateList P inner join SfcGreWorkStationLog L on L.ID=WsLogId where P.SfcId=B.SfcId and PassType=2 and L.WorkStation='02.02.01') and a.status=2 then '完工' 
	when exists(select null from SfcGreOperateList P where P.SfcId=B.SfcId and Operate='FirstComp') and a.status=2 then '量产中' 
	when exists(select null from SfcGreOperateList P where P.SfcId=B.SfcId and Operate='Prepaired') and a.status=2 then '首件生产' 
	when exists(select null from SfcGreOperateList P inner join SfcGreWorkStationLog L on L.ID=WsLogId where P.SfcId=B.SfcId and PassType=1 and L.WorkStation='02.02.01' ) and a.status=2 then '准备中' 
	else '' end '状态', 
	A.Product '产品编号',M.Fname '产品名称',M.Fmodel '规格', 
	substring(convert(nvarchar(16),V.PreBeginTime,20),6,11) '计划开始时间', 
	substring(convert(nvarchar(16),V.PreEndTime,20),6,11) '计划结束时间', 
	Convert(float, A.FQty) '预计产量', 
	case when a.status in(4,5) or (exists(select null from SfcGreOperateList P inner join SfcGreWorkStationLog L on L.ID=WsLogId where P.SfcId=B.SfcId and PassType=2 and L.WorkStation='02.02.01') and a.status=2)  then Convert(float, A.FQty) else isnull(CQ,0) end '完成数量' 
	from SfcExpScheduleView V 
	left join SfcExpSchedule A on V.ScheID=A.ID 
	left join SfcGreNoList B on A.ID=B.ScheduleId 
	left join MdcDatMaterial M on M.Fcode=A.Product 
	left join (select L.ScheduleId,count(*) CQ from SmtGreCountScanData S 
	inner join SfcGreNoList L on S.SfcId=L.SfcId where Scanner='ZY0017' and Barcode<>'ER'  group by ScheduleId) C on C.ScheduleId=A.ID 
	where V.Area='02.02' AND  A.Status<>3 
	AND exists(select null from SfcGreOperateList P where P.SfcId=B.SfcId and Operate='FirstComp') and a.status=2
	and convert(nvarchar(10),V.PreBeginTime,120)!=convert(nvarchar(10),Getdate(),120)
) a
where a.状态 in ('量产中','首件生产')
union ALL
select distinct A.Fcode '计划任务单号', 
case when A.Status=0 then '新建' when A.Status=1 then '保留' when A.Status=4 then '完成' when A.Status=5 then '指定结束' 
when not exists(select null from SfcGreOperateList P inner join SfcGreWorkStationLog L on L.ID=WsLogId where P.SfcId=B.SfcId and PassType=1 and L.WorkStation='02.02.01') and a.status=2 then '未开工' 
when exists(select null from SfcGreOperateList P inner join SfcGreWorkStationLog L on L.ID=WsLogId where P.SfcId=B.SfcId and PassType=2 and L.WorkStation='02.02.01') and a.status=2 then '完工' 
when exists(select null from SfcGreOperateList P where P.SfcId=B.SfcId and Operate='FirstComp') and a.status=2 then '量产中' 
when exists(select null from SfcGreOperateList P where P.SfcId=B.SfcId and Operate='Prepaired') and a.status=2 then '首件生产' 
when exists(select null from SfcGreOperateList P inner join SfcGreWorkStationLog L on L.ID=WsLogId where P.SfcId=B.SfcId and PassType=1 and L.WorkStation='02.02.01' ) and a.status=2 then '准备中' 
else '' end '状态', 
A.Product '产品编号',M.Fname '产品名称',M.Fmodel '规格', 
substring(convert(nvarchar(16),V.PreBeginTime,20),6,11) '计划开始时间', 
substring(convert(nvarchar(16),V.PreEndTime,20),6,11) '计划结束时间', 
Convert(float, A.FQty) '预计产量', 
case when a.status in(4,5) or (exists(select null from SfcGreOperateList P inner join SfcGreWorkStationLog L on L.ID=WsLogId where P.SfcId=B.SfcId and PassType=2 and L.WorkStation='02.02.01') and a.status=2)  then Convert(float, A.FQty) else isnull(CQ,0) end '完成数量' 
from SfcExpScheduleView V 
left join SfcExpSchedule A on V.ScheID=A.ID 
left join SfcGreNoList B on A.ID=B.ScheduleId 
left join MdcDatMaterial M on M.Fcode=A.Product 
left join (select L.ScheduleId,count(*) CQ from SmtGreCountScanData S 
inner join SfcGreNoList L on S.SfcId=L.SfcId where Scanner='ZY0017' and Barcode<>'ER'  group by ScheduleId) C on C.ScheduleId=A.ID 
where V.Area='02.02' AND  A.Status<>3 
and convert(nvarchar(10),V.PreBeginTime,120)=convert(nvarchar(10),Getdate(),120)
order by '计划开始时间'
");

            //连接数据库，绑定表1数据
            SqlConnection conn = new SqlConnection("server=192.168.1.6;database=Exd2CZMestest;uid=sa;pwd=flybarrier");
            SqlDataAdapter sda = new SqlDataAdapter(sql, conn);
            DataSet ds = new DataSet();
            sda.Fill(ds);
            dataGridView1.DataSource = ds.Tables[0];
            //取消选中第一行
            dataGridView1.Rows[0].Selected = false;
            //绑定表2数据
            SqlDataAdapter sda1 = new SqlDataAdapter(sql0, conn);
            DataSet ds1 = new DataSet();
            sda1.Fill(ds1);
            dataGridView2.DataSource = ds1.Tables[0];
            //取消选中第一行
            dataGridView2.Rows[0].Selected = false;
        }

        //状态发生变化，整行字体颜色改变
        private void dataGridView1_CellPainting(object sender, DataGridViewCellPaintingEventArgs e)
        {
            {
                if (e.Value != null && e.RowIndex > -1)
                {
                    DataGridViewRow dgr = dataGridView1.Rows[e.RowIndex];
                    if (e.ColumnIndex == 1)
                    {
                        string strValue = e.Value.ToString();
                        switch (strValue)
                        {
                            case "量产中":
                                this.dataGridView1.Rows[e.RowIndex].DefaultCellStyle.ForeColor = Color.Red;
                                break;
                            case "首件生产":
                                this.dataGridView1.Rows[e.RowIndex].DefaultCellStyle.ForeColor = Color.Red;
                                break;
                            case "未开工":
                                this.dataGridView1.Rows[e.RowIndex].DefaultCellStyle.ForeColor = Color.Blue;
                                break;
                            case "完工":
                                this.dataGridView1.Rows[e.RowIndex].DefaultCellStyle.ForeColor = Color.Green;
                                break;
                            case "准备中":
                                this.dataGridView1.Rows[e.RowIndex].DefaultCellStyle.ForeColor = Color.Red;
                                break;
                        }
                    }
                }
            }
        }

        //状态发生变化，整行字体颜色改变
        private void dataGridView2_CellPainting(object sender, DataGridViewCellPaintingEventArgs e)
        {
            {
                if (e.Value != null && e.RowIndex > -1)
                {
                    DataGridViewRow dgr = dataGridView1.Rows[e.RowIndex];
                    if (e.ColumnIndex == 1)
                    {
                        string strValue = e.Value.ToString();
                        switch (strValue)
                        {
                            case "量产中":
                                this.dataGridView2.Rows[e.RowIndex].DefaultCellStyle.ForeColor = Color.Red;
                                break;
                            case "首件生产":
                                this.dataGridView2.Rows[e.RowIndex].DefaultCellStyle.ForeColor = Color.Red;
                                break;
                            case "未开工":
                                this.dataGridView2.Rows[e.RowIndex].DefaultCellStyle.ForeColor = Color.Blue;
                                break;
                            case "完工":
                                this.dataGridView2.Rows[e.RowIndex].DefaultCellStyle.ForeColor = Color.Green;
                                break;
                            case "准备中":
                                this.dataGridView2.Rows[e.RowIndex].DefaultCellStyle.ForeColor = Color.Red;
                                break;
                        }
                    }
                }
            }
        }

        //DataGridView1行高、单元格字体设置
        private void dataGridView1_CellFormatting(object sender, DataGridViewCellFormattingEventArgs e)
        {
          
             int r = 0;
             if (e.RowIndex < 0)
             {
                 r = 0;
             }
             else    
             {
                 r = e.RowIndex;
             }
             dataGridView1.Rows[r].MinimumHeight = 89;
             dataGridView1.Rows[r].Cells[e.ColumnIndex].Style.Font = new Font("微软雅黑", 30);
             if (r % 2 == 0)
             {
                 dataGridView1.Rows[r].DefaultCellStyle.BackColor = Color.LightSkyBlue;
             }
             else
             {
                 dataGridView1.Rows[r].DefaultCellStyle.BackColor = Color.LightYellow;
             }            
        }

        //DataGridView2行高、单元格字体设置
        private void dataGridView2_CellFormatting(object sender, DataGridViewCellFormattingEventArgs e)
        {
            int r = 0;
            if (e.RowIndex < 0)
            {
                r = 0;
            }
            else
            {
                r = e.RowIndex;
            }
            dataGridView2.Rows[r].MinimumHeight = 89;
            dataGridView2.Rows[r].Cells[e.ColumnIndex].Style.Font = new Font("微软雅黑", 30);
            if (r % 2 == 0)
            {
                dataGridView2.Rows[r].DefaultCellStyle.BackColor = Color.LightSkyBlue;
            }
            else
            {
                dataGridView2.Rows[r].DefaultCellStyle.BackColor = Color.LightYellow;
            }
        }
        //按Esc键退出
        private void Form1_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == System.Windows.Forms.Keys.Escape)
            {
                this.Close();
            }
        }    
    }
}
