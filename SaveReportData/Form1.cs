using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.SqlClient;
using System.Data.OracleClient;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Net.NetworkInformation;
using System.Net;
using System.Threading;
using System.Text.RegularExpressions;
using System.Globalization;
using System.IO;
using System.Net.Mail;
using System.Net.Mime;
using Microsoft.Win32;
using System.Net.Sockets;
using System.Data.Common;
using MySql.Data.MySqlClient;

namespace SaveReportData
{
    public partial class Form1 : Form
    {

        public string strdbisc = "Data Source=(DESCRIPTION =(ADDRESS_LIST =(ADDRESS =(PROTOCOL = TCP)(HOST =10.134.171.33)(PORT = 1527)))(CONNECT_DATA = (SERVICE_NAME = ISCODB)));User Id=PC;Password=iscpc169!!!!";
        public string strToad = "Data Source=(DESCRIPTION =(ADDRESS_LIST =(ADDRESS =(PROTOCOL = TCP)(HOST =10.224.92.136)(PORT = 1527)))(CONNECT_DATA = (SERVICE_NAME = VNSFC)));User Id=HOGAN;Password=sfcsystem2014#!";
        public string strToad1 = "Data Source=(DESCRIPTION =(ADDRESS_LIST =(ADDRESS =(PROTOCOL = TCP)(HOST =10.224.134.59)(PORT = 1527)))(CONNECT_DATA = (SERVICE_NAME = sfcodb)));User Id=sfis1;Password=hift2013!";
        public string strToad2 = "Data Source=(DESCRIPTION =(ADDRESS_LIST =(ADDRESS =(PROTOCOL = TCP)(HOST =10.224.81.34)(PORT = 1521)))(CONNECT_DATA = (SERVICE_NAME = vnsfc)));User Id=sfis1;Password=vnsfis2014#!";
        public string strToad3 = "Data Source=(DESCRIPTION =(ADDRESS_LIST =(ADDRESS =(PROTOCOL = TCP)(HOST =10.224.81.41)(PORT = 1521)))(CONNECT_DATA = (SERVICE_NAME = vncpeap)));User Id=ap2;Password=NSDAP2LOGPD0522";
        public string strToad5 = "Data Source=(DESCRIPTION =(ADDRESS_LIST =(ADDRESS =(PROTOCOL = TCP)(HOST =10.225.35.9)(PORT = 1521)))(CONNECT_DATA = (SERVICE_NAME = vnap)));User Id=ap2;Password=NSDAP2LOGPD0522";
        public string strToad6 = "Data Source=(DESCRIPTION =(ADDRESS_LIST =(ADDRESS =(PROTOCOL = TCP)(HOST =10.224.81.31)(PORT = 1521)))(CONNECT_DATA = (SERVICE_NAME = vnap)));User Id=ap2;Password=NSDAP2LOGPD0522";
        public string strToad7 = "Data Source=(DESCRIPTION =(ADDRESS_LIST =(ADDRESS =(PROTOCOL = TCP)(HOST =10.224.81.33)(PORT = 1521)))(CONNECT_DATA = (SERVICE_NAME = VNSFC)));User Id=GOLDBERG;Password=sfcgold168$@";
        public string strSql1 = "server=10.224.81.131,3000;uid=sa;database=DUTY_HR_INFO;password=foxconn168!!;";
        public string strSql2 = "server=10.224.81.131,3000;uid=sa;database=SmoAlert;password=foxconn168!!;";
        public string strSql3 = "server=10.224.81.131,3000;uid=sa;database=OWNCLOUD_DATA;password=foxconn168!!;";
        public string strSql4 = "server=10.224.81.131,3000;uid=sa;database=TABLEAU_DB;password=foxconn168!!;";
        public string strSql5 = "server=10.224.81.131,3000;uid=sa;database=PingIP;password=foxconn168!!;";
        public string strCn = "server=10.224.81.131,3000;uid=sa;database=findip;password=foxconn168!!;";
        public string strMysql = "Server=10.224.69.111;Database=humanstreaming;port=3306;User Id=root;password=foxconn168!!;Charset=utf8";

        public bool issend = true;
        public bool issend2 = true;
        public bool issend3 = true;
        public bool issend4 = true;

        public string listexception = "";
        public Form1()
        {
            InitializeComponent();
            RegisterInStartup();
        }

        [Obsolete]
        public void scannerdata( string nametable)
        {

            using (OracleConnection con = new OracleConnection(strdbisc))
            {
                int count = 0;
                try
                {
                    con.Open();

                    string sql = "SELECT * FROM " + nametable ;

                    OracleCommand cmd = new OracleCommand(sql, con);

                    OracleDataReader dr = cmd.ExecuteReader();
                    if (dr.HasRows)
                    {
                        while (dr.Read())
                        {
                            string check = dr[1].ToString().Replace("-", ".");
                            if (listexception.Contains(check))
                            {
                                Console.WriteLine("contains " + dr[1].ToString());
                            }
                            else
                            {
                                count++;
                                Console.WriteLine("not contains");
                                string scanagaintime = "2020/08/10 13:46:44";
                                string datetimenow2 = DateTime.Now.ToString("yyyy/MM/dd HH:mm:ss");
                                insertdatasql(dr[0].ToString(), dr[1].ToString(), dr[2].ToString(), dr[3].ToString(), dr[4].ToString(), dr[5].ToString(), dr[6].ToString(), dr[7].ToString(), datetimenow2, nametable);
                                //The system cannot find the file specified.
                            }

                        }
                    }
                    // close and dispose the objects
                    dr.Close();
                    dr.Dispose();
                    cmd.Dispose();
                    con.Close();
                    con.Dispose();

                }
                catch (OracleException ex) // catches only Oracle errors
                {
                    MessageBox.Show("The database is unavailable. "+ex.Message);
                }

                string datetimenow = DateTime.Now.ToString("yyyy/MM/dd HH:mm:ss");

                if (nametable == "VW_NO_INSTALL_SYMANTEC")
                {
                    datetimenow = datetimenow + " | count:  " + count;
                    textBox1.Text = datetimenow.ToString();
                }
                else
                if (nametable == "VW_JOIN_DOMAIN")
                {
                    datetimenow = datetimenow + " | count:  " + count;
                    textBox2.Text = datetimenow.ToString();
                }
                else
                if (nametable == "VW_UPDATE_VERSION_SYMANTEC")
                {
                    datetimenow = datetimenow + " | count:  " + count;
                    textBox3.Text = datetimenow.ToString();
                }
                else
                if (nametable == "VW_OPEN_PORT_DOCNO")
                {
                    datetimenow = datetimenow + " | count:  " + count;
                    textBox5.Text = datetimenow.ToString();
                }
                else
                if (nametable == "VW_USB_AVAI_UNAVAI")
                {
                    datetimenow = datetimenow + " | count:  " + count;
                    textBox7.Text = datetimenow.ToString();
                }
                else
                if (nametable == "VW_USE_PROXY")
                {
                    datetimenow = datetimenow + " | count:  " + count;
                    textBox8.Text = datetimenow.ToString();
                }
                else
                if (nametable == "VW_USE_CD")
                {
                    datetimenow = datetimenow + " | count:  " + count;
                    textBox9.Text = datetimenow.ToString();
                }
                else
                if (nametable == "VW_CONNECT_INTERNET")
                {
                    datetimenow = datetimenow + " | count:  " + count;
                    textBox10.Text = datetimenow.ToString();
                }
                else
                if (nametable == "VW_DEFINE_SOFWARE")
                {
                    datetimenow = datetimenow + " | count:  " + count;
                    textBox6.Text = datetimenow.ToString();
                }
                else
                if (nametable == "vw_server_capacity")
                {
                    datetimenow = datetimenow + " | count:  " + count;
                    textBox27.Text = datetimenow.ToString();
                }
                else
                if (nametable == "VW_DEFINE_BLUETOOL")
                {
                    datetimenow = datetimenow + " | count:  " + count;
                    textBox28.Text = datetimenow.ToString();
                }
                else
                if (nametable == "VW_KERNEL_VERSION")
                {
                    datetimenow = datetimenow + " | count:  " + count;
                    textBox29.Text = datetimenow.ToString();
                }

                Console.WriteLine("Scanner " + nametable +  "Count :" +count);
            }
        }

        public void scannerdataTREND_LINE_PACKAGE(string query ,string TableName)
        {

            using (SqlConnection con = new SqlConnection(strCn))
            {
                try
                {
                    con.Open();
                    int count = 0;
                    string sql = query.Trim();

                    SqlCommand cmd = new SqlCommand(sql, con);

                    SqlDataReader dr = cmd.ExecuteReader();
                    if (dr.HasRows)
                    {
                        while (dr.Read())
                        {
                            count++;
                            string scanagaintime = "2020/08/10 13:46:44";
                            string datetimenow2 = DateTime.Now.ToString("yyyy/MM/dd HH:mm:ss");
                            insertdataTREND_LINE_PACKAGE(dr[1].ToString(), Int32.Parse(dr[0].ToString()), dr[2].ToString(),TableName.Trim());
                        }
                    }
                    // close and dispose the objects
                    dr.Close();
                    dr.Dispose();
                    cmd.Dispose();
                    con.Close();
                    con.Dispose();

                    string datetimenow = DateTime.Now.ToString("yyyy/MM/dd HH:mm:ss");
                    textBox11.Text = datetimenow.ToString()+" | Count :"+count;

                }
                catch (OracleException ex) // catches only Oracle errors
                {
                    MessageBox.Show("The database is unavailable."+ex.Message);
                }        

            }
        }

        public void scannerdataTREND_LINE_PACKAGE_DETAIL(string query,string TableName)
        {

            using (OracleConnection con = new OracleConnection(strdbisc))
            {
                try
                {
                    con.Open();
                    int count = 0;
                    string sql = query.Trim();
                    OracleCommand cmd = new OracleCommand(sql, con);
                    OracleDataReader dr = cmd.ExecuteReader();
                    if (dr.HasRows)
                    {
                        while (dr.Read())
                        {
                            string scanagaintime = "2020/08/10 13:46:44";
                            string datetimenow2 = DateTime.Now.ToString("yyyy/MM/dd HH:mm:ss");
                            count++;
                            insertdataTREND_LINE_PACKAGE_DETAIL(dr[0].ToString(), dr[1].ToString(), dr[2].ToString(), dr[3].ToString(), dr[4].ToString(), dr[5].ToString(), dr[6].ToString(), dr[7].ToString(), datetimenow2, TableName.Trim());
                        }
                    }
                    // close and dispose the objects
                    dr.Close();
                    dr.Dispose();
                    cmd.Dispose();
                    con.Close();
                    con.Dispose();

                    string datetimenow = DateTime.Now.ToString("yyyy/MM/dd HH:mm:ss");
                    textBox12.Text = datetimenow.ToString()+" | Count :"+count;
                }
                catch (OracleException ex) // catches only Oracle errors
                {
                    MessageBox.Show("The database is unavailable."+ex.Message);
                }

            }
        }

        public void scannerVW_CHROME_VERSION()
        {

            using (OracleConnection con = new OracleConnection(strdbisc))
            {
                try
                {
                    con.Open();
                    int count = 0;
                    string sql = "select PCNAME,USERNAME,CUST_KP_NO,EMP_NO,INPUT_DATE,IPS,HR,CITY_LOMO from VW_CHROME_VERSION";

                    OracleCommand cmd = new OracleCommand(sql, con);

                    OracleDataReader dr = cmd.ExecuteReader();
                    if (dr.HasRows)
                    {
                        while (dr.Read())
                        {
                            count++;
                            string scanagaintime = "2020/08/10 13:46:44";
                            string datetimenow2 = DateTime.Now.ToString("yyyy/MM/dd HH:mm:ss");
                            insertDataVW_CHROME_VERSION(dr[0].ToString(), dr[1].ToString(), dr[2].ToString(), dr[3].ToString(), dr[4].ToString(), dr[5].ToString(), dr[6].ToString(), dr[7].ToString(), datetimenow2, "dbo.VW_CHROME_VERSION");

                        }
                    }
                    // close and dispose the objects
                    dr.Close();
                    dr.Dispose();
                    cmd.Dispose();
                    con.Close();
                    con.Dispose();
                    string datetimenow = DateTime.Now.ToString("yyyy/MM/dd HH:mm:ss");
                    textBox22.Text = datetimenow.ToString() + " | Count :"+count;

                }
                catch (OracleException ex) // catches only Oracle errors
                {
                    MessageBox.Show("The database is unavailable."+ex.Message);
                }

            }
        }

        public void scannerVW_GZ_PC()
        {

            using (OracleConnection con = new OracleConnection(strdbisc))
            {
                try
                {
                    DeleteTable("VW_GZ_PC", strSql5);
                    con.Open();
                    int count = 0;
                    string sql = "select PCNAME,USERNAME,CUST_KP_NO,EMP_NO,INPUT_DATE,IPS,HR,CITY_LOMO from VW_GZ_PC";

                    OracleCommand cmd = new OracleCommand(sql, con);

                    OracleDataReader dr = cmd.ExecuteReader();
                    if (dr.HasRows)
                    {
                        while (dr.Read())
                        {
                            count++;
                            string scanagaintime = "2020/08/10 13:46:44";
                            string datetimenow2 = DateTime.Now.ToString("yyyy/MM/dd HH:mm:ss");
                            insertDataVW_CHROME_VERSION(dr[0].ToString(), dr[1].ToString(), dr[2].ToString(), dr[3].ToString(), dr[4].ToString(), dr[5].ToString(), dr[6].ToString(), dr[7].ToString(), datetimenow2, "VW_GZ_PC");

                        }
                    }
                    // close and dispose the objects
                    dr.Close();
                    dr.Dispose();
                    cmd.Dispose();
                    con.Close();
                    con.Dispose();

                    string datetimenow = DateTime.Now.ToString("yyyy/MM/dd HH:mm:ss");
                    textBox23.Text = datetimenow.ToString()+" Count :"+count;

                }
                catch (OracleException ex) // catches only Oracle errors
                {
                    MessageBox.Show("The database is unavailable.");
                } 

            }
        }

        public void scannerVW_HT_PC()
        {

            using (OracleConnection con = new OracleConnection(strdbisc))
            {
                try
                {
                    DeleteTable("VW_HT_PC", strSql5);
                    con.Open();
                    int count = 0;
                    string sql = "select PCNAME,USERNAME,CUST_KP_NO,EMP_NO,INPUT_DATE,IPS,HR,CITY_LOMO from VW_HT_PC ";

                    OracleCommand cmd = new OracleCommand(sql, con);

                    OracleDataReader dr = cmd.ExecuteReader();
                    if (dr.HasRows)
                    {
                        while (dr.Read())
                        {
                            count++;
                            string scanagaintime = "2020/08/10 13:46:44";
                            string datetimenow2 = DateTime.Now.ToString("yyyy/MM/dd HH:mm:ss");
                            insertDataVW_CHROME_VERSION(dr[0].ToString(), dr[1].ToString(), dr[2].ToString(), dr[3].ToString(), dr[4].ToString(), dr[5].ToString(), dr[6].ToString(), dr[7].ToString(), datetimenow2, "VW_HT_PC ");

                        }
                    }
                    // close and dispose the objects
                    dr.Close();
                    dr.Dispose();
                    cmd.Dispose();
                    con.Close();
                    con.Dispose();

                    string datetimenow = DateTime.Now.ToString("yyyy/MM/dd HH:mm:ss");
                    textBox24.Text = datetimenow.ToString() +" | Count: "+count;

                }
                catch (OracleException ex) // catches only Oracle errors
                {
                    MessageBox.Show("The database is unavailable.");
                }

            }
        }
        //sanner data and delete by day
        public void DeleteVW_ROOT_PACKAGE()
        {
            string sql_delete = "delete from [dbo].[VW_ROOT_PACKAGE]";

            SqlConnection con = new SqlConnection(strCn);
            SqlCommand cmd = new SqlCommand(sql_delete, con);
            try
            {
                con.Open();
                cmd.ExecuteNonQuery();
            }
            catch (SqlException ex)
            {
            }
            finally
            {
                con.Close();
            }
        }
        public void scannerdataVW_ROOT_PACKAGE()
        {
            DeleteVW_ROOT_PACKAGE();
            using (OracleConnection con = new OracleConnection(strdbisc))
            {
                try
                {
                    con.Open();
                    int count = 0;
                    string sql = "SELECT * FROM WINDOW_UPDATE_PACKAGE_win10 WHERE pcname = '00-15-5D-28-FB-04' union all SELECT* FROM WINDOW_UPDATE_PACKAGE_server16 WHERE pcname = 'F4-03-43-55-81-4D' union all SELECT* FROM WINDOW_UPDATE_PACKAGE_server12 WHERE pcname = '98-F2-B3-3F-F1-84' union all SELECT* FROM WINDOW_UPDATE_PACKAGE_server19 WHERE pcname = '00-15-5D-45-5C-06'";

                    OracleCommand cmd = new OracleCommand(sql, con);

                    OracleDataReader dr = cmd.ExecuteReader();
                    if (dr.HasRows)
                    {
                        while (dr.Read()) 
                        {
                            count++;
                            string scanagaintime = "2020/08/10 13:46:44";
                            string datetimenow2 = DateTime.Now.ToString("yyyy/MM/dd HH:mm:ss");
                            insertdataVW_ROOT_PACKAGE(dr[0].ToString(), dr[1].ToString(), dr[2].ToString(), dr[3].ToString(), dr[4].ToString(), datetimenow2);

                        }
                    }
                    // close and dispose the objects
                    dr.Close();
                    dr.Dispose();
                    cmd.Dispose();
                    con.Close();
                    con.Dispose();
                    string datetimenow = DateTime.Now.ToString("yyyy/MM/dd HH:mm:ss");
                    textBox18.Text = datetimenow.ToString()+" | Count :"+count;

                }
                catch (OracleException ex) // catches only Oracle errors
                {
                    MessageBox.Show("The database is unavailable.");
                }
              

            }
        }
        //public void scannerdataport(string nametable)
        //{
        //    int count = 0;
        //    using (OracleConnection con = new OracleConnection(strdbisc))
        //    {
        //        try
        //        {
        //            con.Open()
                    
        //            string sql = "SELECT * FROM VW_OPEN_PORT_DOCNO";
        //            OracleCommand cmd = new OracleCommand(sql, con);
        //            OracleDataReader dr = cmd.ExecuteReader();
        //            if (dr.HasRows)
        //            {
        //                while (dr.Read())
        //                {
        //                    count++;
        //                    string check = dr[1].ToString().Replace("-", ".");
        //                    if (listexception.Contains(check))
        //                    {
        //                        Console.WriteLine("contains " + dr[1].ToString());
        //                    }
        //                    else
        //                    {
        //                        Console.WriteLine("not contains");
        //                        string scanagaintime = "2020/08/10 13:46:44";
        //                        string datetimenow2 = DateTime.Now.ToString("yyyy/MM/dd HH:mm:ss");
        //                        insertdatadetectportsql(dr[0].ToString(), dr[1].ToString(), dr[2].ToString(), dr[3].ToString(), dr[4].ToString(), dr[5].ToString(), dr[6].ToString(), dr[7].ToString(), datetimenow2, nametable, dr[8].ToString(), dr[9].ToString(), dr[10].ToString());
        //                    }

        //                }
        //            }
        //            // close and dispose the objects
        //            dr.Close();
        //            dr.Dispose();
        //            cmd.Dispose();
        //            con.Close();
        //            con.Dispose();

        //        }
        //        catch (OracleException ex) // catches only Oracle errors
        //        {
        //            MessageBox.Show("The database is unavailable.");
        //        }
        //        string datetimenow = DateTime.Now.ToString("yyyy/MM/dd HH:mm:ss");
        //        if (nametable == "VW_OPEN_PORT_DOCNO")
        //        {
        //            textBox5.Text = datetimenow.ToString()+" | Count :"+count;
        //        }
        //        Console.WriteLine("Scanner " + nametable + " OKEEEEEEEEEEEE");
        //    }
        //}
        public void scannerUpdateWindowStatusdata(string table)
        {

            using (OracleConnection con = new OracleConnection(strdbisc))
            {
                try
                {
                    con.Open();
                    int count = 0;
                    string sql = "SELECT * FROM "+table;
                    OracleCommand cmd = new OracleCommand(sql, con);
                    OracleDataReader dr = cmd.ExecuteReader();
                    if (dr.HasRows)
                    {
                        while (dr.Read())
                        {
                            string check = dr[1].ToString().Replace("-", ".");
                            if (listexception.Contains(check))
                            {
                                Console.WriteLine("contains");
                            }
                            else
                            {
                                Console.WriteLine("not contains");
                                string scanagaintime = "2020/08/10 13:46:44";
                                string datetimenow2 = DateTime.Now.ToString("yyyy/MM/dd HH:mm:ss");
                                count++;
                                insertdatasqlupdatewindows(dr[0].ToString(), dr[1].ToString(), dr[2].ToString(), dr[3].ToString(), dr[4].ToString(), dr[5].ToString(), dr[6].ToString(), datetimenow2, "VW_UPDATE_WINDOW_STATUS_ALL");
                            }              
                        }
                    }
                    // close and dispose the objects
                    dr.Close();
                    dr.Dispose();
                    cmd.Dispose();
                    con.Close();
                    con.Dispose();

                    string datetimenow = DateTime.Now.ToString("yyyy/MM/dd HH:mm:ss");
                    textBox4.Text = datetimenow.ToString() + " | Count :"+count;
                }
                catch (OracleException ex) // catches only Oracle errors
                {
                    MessageBox.Show("The database is unavailable."+ex.Message);
                }
               

            }
        }
        //scandate C_PC_CONTROL_T 2020/09/11
        public void scannerC_PC_CONTROL_T()
        {
            using (OracleConnection con = new OracleConnection(strdbisc))
            {
                try
                {
                    con.Open();
                    int count = 0;
                    string sql = "select * from MES1.C_PC_CONTROL_T WHERE hr='PCNAME' ORDER BY INPUT_DATE ASC";
                    OracleCommand cmd = new OracleCommand(sql, con);

                    OracleDataReader dr = cmd.ExecuteReader();
                    if (dr.HasRows)
                    {
                        while (dr.Read())
                        {
                            count++;
                            string check = dr[1].ToString().Replace("-", ".");
                            if (listexception.Contains(check))
                            {
                                Console.WriteLine("contains " + dr[1].ToString());
                            }
                            else
                            {
                                Console.WriteLine("not contains");
                                string scanagaintime = "2020/08/10 13:46:44";
                                string datetimenow2 = DateTime.Now.ToString("yyyy/MM/dd HH:mm:ss");
                                insertDataPcControl(dr[0].ToString(), dr[1].ToString(), dr[2].ToString(), dr[3].ToString(), dr[4].ToString(), datetimenow2.ToString());
                            }

                        }
                    }
                    // close and dispose the objects
                    dr.Close();
                    dr.Dispose();
                    cmd.Dispose();
                    con.Close();
                    con.Dispose();
                    string datetimenow = DateTime.Now.ToString("yyyy/MM/dd HH:mm:ss");
                    textBox14.Text = datetimenow.ToString()+" Count : "+count;

                }
                catch (OracleException ex) // catches only Oracle errors
                {
                    MessageBox.Show("The database is unavailable.");
                }
               
            }
        }
        //scanner data for toad oracle chi than
        //update sql date 9/3/2020
        //Scandate 
        public void scannerC_PC_NOTOPEN()
        {
            //scandate 15/09/2020
            using (OracleConnection con = new OracleConnection(strdbisc))
            {
                int count=0;
                try
                {
                    con.Open();
                    string sql = "SELECT MACID,IP,RANGEIP,EMP_NO,INPUT_DATE,PCNAME,HR,LOCATION  FROM VW_PCNAME_NOTOPEN";
                    OracleCommand cmd = new OracleCommand(sql, con);

                    OracleDataReader dr = cmd.ExecuteReader();
                     
                    if (dr.HasRows)
                    {

                        while (dr.Read())
                        {
                            count=count+1;
                            Console.WriteLine("not contains");
                            string scanagaintime = "2020/08/10 13:46:44";
                            string datetimenow2 = DateTime.Now.ToString("yyyy/MM/dd HH:mm:ss");
                            insertDataPcNotOpen(dr[0].ToString(), dr[1].ToString(), dr[2].ToString(), dr[3].ToString(), dr[4].ToString(), dr[5].ToString(), dr[6].ToString(), dr[7].ToString(), datetimenow2.ToString());
                        }
                    }
                    // close and dispose the objects
                    dr.Close();
                    dr.Dispose();
                    cmd.Dispose();
                    con.Close();
                    con.Dispose();
                    string datetimenow = DateTime.Now.ToString("yyyy/MM/dd HH:mm:ss");
                    textBox16.Text = datetimenow.ToString()+" | Count :"+count;
                }
                catch (OracleException ex) // catches only Oracle errors
                {
                    MessageBox.Show("The database is unavailable."+ex.Message);
                }
 
            }
        }
        // Scanner date NTP
        public void scannerDataNTP()
        {
            //scandate 15/09/2020
            using (OracleConnection con = new OracleConnection(strdbisc))
            {
                try
                {
                    con.Open();
                    int count = 0;
                    string sql = "select a.PCNAME,a.USERNAME,a.NTP,a.NOW,a.RANGETIME,b.emp_no,b.city_lomo,b.mm,b.cust_kp_no,a.INPUT_dATE from (SELECT PCNAME, USERNAME, NTP, NOW, INPUT_dATE, RANGETIME, REPLACE(SUBSTR(USERNAME, 1, INSTR(USERNAME, '-', 1, 3) - 1), '-', '') as st,  SUBSTR(username, INSTR(username, '-', 1, 3) + 1)hh FROM VW_NTP_CLENT_TIME) a, (select * from range_ip) b where b.mm(+) = a.st";
                    OracleCommand cmd = new OracleCommand(sql, con);

                    OracleDataReader dr = cmd.ExecuteReader();

                    if (dr.HasRows)
                    {
                        while (dr.Read())
                        {
                            count++;
                            Console.WriteLine("not contains");
                            string scanagaintime = "2020/08/10 13:46:44";
                            string datetimenow2 = DateTime.Now.ToString("yyyy/MM/dd HH:mm:ss");
                            insertDataNTP(dr[0].ToString(), dr[1].ToString(), dr[2].ToString(), dr[3].ToString(), dr[4].ToString(), dr[5].ToString(), dr[6].ToString(), dr[7].ToString(), dr[8].ToString(), dr[9].ToString(), datetimenow2.ToString());
                        }
                    }
                    // close and dispose the objects
                    dr.Close();
                    dr.Dispose();
                    cmd.Dispose();
                    con.Close();
                    con.Dispose();

                    string datetimenow = DateTime.Now.ToString("yyyy/MM/dd HH:mm:ss");
                    textBox17.Text = datetimenow.ToString()+" | Count :"+count;

                }
                catch (OracleException ex) // catches only Oracle errors
                {
                    MessageBox.Show("The database is unavailable."+ex.Message);
                }
               
            }
        }
        public void scannerCHECK_DAILY()
        {
            using (OracleConnection con = new OracleConnection(strToad))
            {
                try
                {
                    con.Open();
                    int count = 0;
                    string sql = "select a.mail_id,a.location,a.server_ip,a.database_name,a.PROGRAM_NAME,a.DESCRIBE,a.owner,a.date_dontsend as time,b.emp_name from (SELECT mail_id,location,server_ip,database_name,PROGRAM_NAME,DESCRIBE,owner,to_char(sysdate-1,'MM/DD/YYYY')as date_dontsend  FROM SFIS1.C_MAIL_CONFIG_T WHERE mail_id NOT IN( select DISTINCT MAIL_ID from SFIS1.VW_DONT_SEND_MAIL_YESTERDAY))a, (select  emp_no,emp_name from SFIS1.C_EMP_DESC_T ) b where  a.owner=B.EMP_NO(+)";

                    OracleCommand cmd = new OracleCommand(sql, con);

                    OracleDataReader dr = cmd.ExecuteReader();
                    if (dr.HasRows)
                    {
                        while (dr.Read())
                        {
                            
                            string check = dr[1].ToString().Replace("-", ".");
                            if (listexception.Contains(check))
                            {
                                Console.WriteLine("contains " + dr[1].ToString());
                            }
                            else
                            {
                                Console.WriteLine("not contains");
                                string scanagaintime = "2020/08/10 13:46:44";
                                string datetimenow2 = DateTime.Now.ToString("yyyy/MM/dd HH:mm:ss");
                                count ++;
                                insertDataDailyMail(dr[0].ToString(), dr[1].ToString(), dr[2].ToString(), dr[3].ToString(), dr[4].ToString(), dr[5].ToString(), dr[6].ToString(), dr[7].ToString(), dr[8].ToString(), datetimenow2.ToString());
                                
                            }

                        }
                    }
                    // close and dispose the objects
                    dr.Close();
                    dr.Dispose();
                    cmd.Dispose();
                    con.Close();
                    con.Dispose();
                    string datetimenow = DateTime.Now.ToString("yyyy/MM/dd HH:mm:ss");
                    textBox14.Text = datetimenow.ToString()+"Count :"+ count;
                }
                catch (OracleException ex) // catches only Oracle errors
                {
                    MessageBox.Show("The database is unavailable."+ex.Message);
                }
                
            }
        }
        //scanner data chi than 19/08/2020
        public void scannerCHECK_EMP()
        {

            using (OracleConnection con = new OracleConnection(strToad1))
            {
                try
                {
                    con.Open();

                    string sql = "SELECT emp_no,emp_name,id,amount,check_from,check_to FROM VW_CHECK_VERSION_SUMMARY WHERE amount = (select max(amount) from VW_CHECK_VERSION_SUMMARY)";

                    OracleCommand cmd = new OracleCommand(sql, con);

                    OracleDataReader dr = cmd.ExecuteReader();
                    if (dr.HasRows)
                    {
                        while (dr.Read())
                        {
                            string check = dr[1].ToString().Replace("-", ".");
                            if (listexception.Contains(check))
                            {
                                Console.WriteLine("contains " + dr[1].ToString());
                            }
                            else
                            {
                                Console.WriteLine("not contains");
                                string scanagaintime = "2020/08/10 13:46:44";
                                string datetimenow2 = DateTime.Now.ToString("yyyy/MM/dd HH:mm:ss");

                                insertDataCheckEmp(dr[0].ToString(), dr[1].ToString(), dr[2].ToString(), dr[3].ToString(), dr[4].ToString(), dr[5].ToString(), datetimenow2.ToString());

                            }

                        }
                    }
                    // close and dispose the objects
                    dr.Close();
                    dr.Dispose();
                    cmd.Dispose();
                    con.Close();
                    con.Dispose();

                }
                catch (OracleException ex) // catches only Oracle errors
                {
                    MessageBox.Show("The database is unavailable.");
                }
                
            }
        }
        public void scannerCHECK_EMP1()
        {
            using (SqlConnection con = new SqlConnection(strCn))
            {
                try
                {
                    con.Open();

                    string sql = "select owner emp_no, emp_name ,ID,AMOUNT,check_from,check_to  from VW_sendmail_summary ";

                    SqlCommand cmd = new SqlCommand(sql, con);

                    SqlDataReader dr = cmd.ExecuteReader();
                    if (dr.HasRows)
                    {
                        while (dr.Read())
                        {
                            string check = dr[1].ToString().Replace("-", ".");
                            if (listexception.Contains(check))
                            {
                                Console.WriteLine("contains " + dr[1].ToString());
                            }
                            else
                            {
                                Console.WriteLine("not contains");
                                string scanagaintime = "2020/08/10 13:46:44";
                                string datetimenow2 = DateTime.Now.ToString("yyyy/MM/dd HH:mm:ss");

                                insertDataCheckEmp(dr[0].ToString(), dr[1].ToString(), dr[2].ToString(), dr[3].ToString(), dr[4].ToString(), dr[5].ToString(), datetimenow2.ToString());

                            }

                        }
                    }
                    // close and dispose the objects
                    dr.Close();
                    dr.Dispose();
                    cmd.Dispose();
                    con.Close();
                    con.Dispose();

                }
                catch (OracleException ex) // catches only Oracle errors
                {
                    MessageBox.Show("The database is unavailable.");
                }
               
            }
        }
        public void scannerCHECK_EMP2()
        {
            using (SqlConnection con = new SqlConnection(strSql1))
            {
                try
                {
                    con.Open();

                    string sql = "SELECT F_EMPNO AS emp_no, F_NAME emp_name ,ID ,AMOUNT,check_from,check_to FROM VW_DUTY_SUMMARRY";

                    SqlCommand cmd = new SqlCommand(sql, con);

                    SqlDataReader dr = cmd.ExecuteReader();
                    if (dr.HasRows)
                    {
                        while (dr.Read())
                        {
                            string check = dr[1].ToString().Replace("-", ".");
                            if (listexception.Contains(check))
                            {
                                Console.WriteLine("contains " + dr[1].ToString());
                            }
                            else
                            {
                                Console.WriteLine("not contains");
                                string scanagaintime = "2020/08/10 13:46:44";
                                string datetimenow2 = DateTime.Now.ToString("yyyy/MM/dd HH:mm:ss");

                                insertDataCheckEmp(dr[0].ToString(), dr[1].ToString(), dr[2].ToString(), dr[3].ToString(), dr[4].ToString(), dr[5].ToString(), datetimenow2.ToString());

                            }

                        }
                    }
                    // close and dispose the objects
                    dr.Close();
                    dr.Dispose();
                    cmd.Dispose();
                    con.Close();
                    con.Dispose();

                }
                catch (OracleException ex) // catches only Oracle errors
                {
                    MessageBox.Show("The database is unavailable.");
                }
            }
        }
        public void scannerCHECK_EMP3()
        {
            using (SqlConnection con = new SqlConnection(strSql2))
            {
                try
                {
                    con.Open();

                    string sql = "SELECT emp_no,emp_name,id,amount,GETDATE() as  check_from,getdate()-7 as check_to FROM VW_SMO_SUMMARRY";

                    SqlCommand cmd = new SqlCommand(sql, con);

                    SqlDataReader dr = cmd.ExecuteReader();
                    if (dr.HasRows)
                    {
                        while (dr.Read())
                        {
                            string check = dr[1].ToString().Replace("-", ".");
                            if (listexception.Contains(check))
                            {
                                Console.WriteLine("contains " + dr[1].ToString());
                            }
                            else
                            {
                                Console.WriteLine("not contains");
                                string scanagaintime = "2020/08/10 13:46:44";
                                string datetimenow2 = DateTime.Now.ToString("yyyy/MM/dd HH:mm:ss");

                                insertDataCheckEmp(dr[0].ToString(), dr[1].ToString(), dr[2].ToString(), dr[3].ToString(), dr[4].ToString(), dr[5].ToString(), datetimenow2.ToString());
                            }
                        }
                    }
                    // close and dispose the objects
                    dr.Close();
                    dr.Dispose();
                    cmd.Dispose();
                    con.Close();
                    con.Dispose();
                }
                catch (OracleException ex) // catches only Oracle errors
                {
                    MessageBox.Show("The database is unavailable.");
                }
                string datetimenow = DateTime.Now.ToString("yyyy/MM/dd HH:mm:ss");
                textBox13.Text = datetimenow.ToString();
            }
        }
        //scanner emp 12/09/2020
        public void scannerCHECK_EMP4()
        {

            using (OracleConnection con = new OracleConnection(strToad2))
            {
                try
                {
                    con.Open();

                    string sql = "SELECT emp_no,emp_name,id,amount,check_from,check_to FROM VW_DBHEATH_SUMMARY";

                    OracleCommand cmd = new OracleCommand(sql, con);

                    OracleDataReader dr = cmd.ExecuteReader();
                    if (dr.HasRows)
                    {
                        while (dr.Read())
                        {
                            string datetimenow2 = DateTime.Now.ToString("yyyy/MM/dd HH:mm:ss");
                            insertDataCheckEmp(dr[0].ToString(), dr[1].ToString(), dr[2].ToString(), dr[3].ToString(), dr[4].ToString(), dr[5].ToString(), datetimenow2.ToString());

                        }
                    }
                    // close and dispose the objects
                    dr.Close();
                    dr.Dispose();
                    cmd.Dispose();
                    con.Close();
                    con.Dispose();

                }
                catch (OracleException ex) // catches only Oracle errors
                {
                    MessageBox.Show("The database is unavailable."+ strToad2);
                }

            }
        }
        //scanner emp 14/092020
        //update 18/09/2020 SELECT emp_no,emp_name,id,amount,check_from,check_to FROM VW_DBHEATH_SUMMARY
        public void scannerCHECK_EMP5()
        {

            using (OracleConnection con = new OracleConnection(strToad3))
            {
                try
                {
                    con.Open();

                    string sql = "select 'V0955250' emp_no,'阮文進' as emp_name,'DBheath' as ID ,SUM(count) AS AMOUNT,SYSDATE AS check_from,sysdate-7 check_to  from (select count(*) as count from VW_DBHEATH_TABLEAU_30DAYS where TO_CHAR (time, 'YYYY/MM/DD') >= to_char(sysdate- 7,'YYYY/MM/DD') and check_pr ='Warning')";

                    OracleCommand cmd = new OracleCommand(sql, con);

                   OracleDataReader dr = cmd.ExecuteReader();
                    if (dr.HasRows)
                    {
                        while (dr.Read())
                        {
                            string datetimenow2 = DateTime.Now.ToString("yyyy/MM/dd HH:mm:ss");
                            insertDataCheckEmp(dr[0].ToString(), dr[1].ToString(), dr[2].ToString(), dr[3].ToString(), dr[4].ToString(), dr[5].ToString(), datetimenow2.ToString());
                            
                        }
                    }
                    // close and dispose the objects
                    dr.Close();
                    dr.Dispose();
                    cmd.Dispose();
                    con.Close();
                    con.Dispose();

                }
                catch (OracleException ex) // catches only Oracle errors
                {
                    MessageBox.Show("The database is unavailable.");
                }

            }
        }
        public void scannerCHECK_EMP6()
        {

            using (OracleConnection con = new OracleConnection(strToad6))
            {
                try
                {
                    con.Open();

                    string sql = "select 'V0958954' emp_no,'陳德龍' as emp_name,'DBheath' as ID ,SUM(count) AS AMOUNT,SYSDATE AS check_from,sysdate-7 check_to  from (select count(*) as count from VW_DBHEATH_TABLEAU_30DAYS where TO_CHAR (time, 'YYYY/MM/DD') >= to_char(sysdate- 7,'YYYY/MM/DD') and check_pr ='Warning')";

                    OracleCommand cmd = new OracleCommand(sql, con);

                    OracleDataReader dr = cmd.ExecuteReader();
                    if (dr.HasRows)
                    {
                        while (dr.Read())
                        {
                            string datetimenow2 = DateTime.Now.ToString("yyyy/MM/dd HH:mm:ss");
                            insertDataCheckEmp(dr[0].ToString(), dr[1].ToString(), dr[2].ToString(), dr[3].ToString(), dr[4].ToString(), dr[5].ToString(), datetimenow2.ToString()); 
                        }
                    }
                    // close and dispose the objects
                    dr.Close();
                    dr.Dispose();
                    cmd.Dispose();
                    con.Close();
                    con.Dispose();

                }
                catch (OracleException ex) // catches only Oracle errors
                {
                    MessageBox.Show("The database is unavailable.");
                }

            }
        }
        public void scannerCHECK_EMP7()
        {

            using (OracleConnection con = new OracleConnection(strToad5))
            {
                try
                {
                    con.Open();

                    string sql = "select 'V0923811' emp_no,'武黃鐘' as emp_name,'DBheath' as ID ,SUM(count) AS AMOUNT,SYSDATE AS check_from,sysdate-7 check_to  from (select count(*) as count from VW_DBHEATH_TABLEAU_30DAYS where TO_CHAR (time, 'YYYY/MM/DD') >= to_char(sysdate- 7,'YYYY/MM/DD') and check_pr ='Warning')";

                    OracleCommand cmd = new OracleCommand(sql, con);

                    OracleDataReader dr = cmd.ExecuteReader();
                    if (dr.HasRows)
                    {
                        while (dr.Read())
                        {
                            string datetimenow2 = DateTime.Now.ToString("yyyy/MM/dd HH:mm:ss");
                            insertDataCheckEmp(dr[0].ToString(), dr[1].ToString(), dr[2].ToString(), dr[3].ToString(), dr[4].ToString(), dr[5].ToString(), datetimenow2.ToString());
                            
                        }
                    }
                    // close and dispose the objects
                    dr.Close();
                    dr.Dispose();
                    cmd.Dispose();
                    con.Close();
                    con.Dispose();

                }
                catch (OracleException ex) // catches only Oracle errors
                {
                    MessageBox.Show("The database is unavailable.");
                }

            }
        }
        //update 1/10/2020
        public void scannerCHECK_EMP8()
        {
            using (SqlConnection con = new SqlConnection(strSql3))
            {
                try
                {
                    con.Open();

                    string sql = "select Owner,NameOwner, 'webservice' as ID, count(*) as amount  , GETDATE() as  check_from,getdate()-7 as check_to from   VW_service_summary  group by Owner,NameOwner";

                    SqlCommand cmd = new SqlCommand(sql, con);

                    SqlDataReader dr = cmd.ExecuteReader();
                    if (dr.HasRows)
                    {
                        while (dr.Read())
                        {
                            string check = dr[1].ToString().Replace("-", ".");
                            if (listexception.Contains(check))
                            {
                                Console.WriteLine("contains " + dr[1].ToString());
                            }
                            else
                            {
                                Console.WriteLine("not contains");
                                string scanagaintime = "2020/08/10 13:46:44";
                                string datetimenow2 = DateTime.Now.ToString("yyyy/MM/dd HH:mm:ss");

                                insertDataCheckEmp(dr[0].ToString(), dr[1].ToString(), dr[2].ToString(), dr[3].ToString(), dr[4].ToString(), dr[5].ToString(), datetimenow2.ToString());

                            }

                        }
                    }
                    // close and dispose the objects
                    dr.Close();
                    dr.Dispose();
                    cmd.Dispose();
                    con.Close();
                    con.Dispose();

                }
                catch (OracleException ex) // catches only Oracle errors
                {
                    MessageBox.Show("The database is unavailable.");
                }
                string datetimenow = DateTime.Now.ToString("yyyy/MM/dd HH:mm:ss");
                textBox15.Text = datetimenow.ToString();
            }
        }
        //Scanner mysql 11/07/2020
        public void scannerMysql_emp_esd()
        {
           
            using (MySqlConnection con = new MySqlConnection(strMysql))
            {
                try
                {
                    con.Open();
                    int count = 0;
                    string sql = "SELECT f_empno,f_name,f_departname from humanstreaming.emp_basic_info_new WHERE utc_f_untogroup IS NULL and f_site='桂武工業區'";

                    MySqlCommand cmd = new MySqlCommand(sql, con);  
                    MySqlDataReader dr = cmd.ExecuteReader();

                    DeleteTable("dbo.emp_esd", strSql4);
                    if (dr.HasRows)
                    {
                        while (dr.Read())
                        {
                            count++;
                           string scanagaintime = "2020/08/10 13:46:44";
                           string datetimenow2 = DateTime.Now.ToString("yyyy/MM/dd HH:mm:ss");
                           insertDataMysql(dr[0].ToString(), dr[1].ToString(), dr[2].ToString());
                        }
                    }
                    // close and dispose the objects
                    dr.Close();
                    dr.Dispose();
                    cmd.Dispose();
                    con.Close();
                    con.Dispose();
                    string datetimenow = DateTime.Now.ToString("yyyy/MM/dd HH:mm:ss");
                    textBox21.Text = datetimenow.ToString() + " | Count :"+count;
                }
                catch (MySqlException ex) // catches only Oracle errors
                {
                    MessageBox.Show("The database is unavailable."+ex.Message);
                }
                
            }
        }

        public void scannerMysql_emp_esd_all()
        {
            using (MySqlConnection con = new MySqlConnection(strMysql))
            {
                int count = 0;
                try
                {
                    con.Open();
                    string sql = "SELECT * FROM humanstreaming.emp_basic_info_new where utc_f_untogroup is null ";
                    MySqlCommand cmd = new MySqlCommand(sql, con);
                    MySqlDataReader dr = cmd.ExecuteReader();
                    DeleteTable("dbo.emp_esd_all", strSql4);
                    if (dr.HasRows)
                    {
                        while (dr.Read())
                        {
                            count++;
                            string scanagaintime = dr[0].ToString();
                            string datetimenow2 = DateTime.Now.ToString("yyyy/MM/dd HH:mm:ss");
                            if(string.IsNullOrEmpty(dr[0].ToString()) || dr[0].ToString() == "")
                            {
                                continue;
                            }
                            else
                            {
                                insertDataMysql_emp_esd_all(dr[0].ToString(), dr[1].ToString(), dr[2].ToString(), dr[3].ToString(), dr[4].ToString(), dr[5].ToString(), dr[6].ToString(), dr[7].ToString(), dr[8].ToString(), dr[9].ToString(), dr[10].ToString(), dr[11].ToString(), dr[12].ToString(), dr[13].ToString(), dr[14].ToString(), dr[15].ToString());
                            }
                            Console.WriteLine(scanagaintime);
                        }
                    }
                    // close and dispose the objects
                    dr.Close();
                    dr.Dispose();
                    cmd.Dispose();
                    con.Close();
                    con.Dispose();

                }
                catch (MySqlException ex) // catches only Oracle errors
                {
                    MessageBox.Show("The database is unavailable." + ex.Message);
                }
                string datetimenow = DateTime.Now.ToString("yyyy/MM/dd HH:mm:ss");
                datetimenow = datetimenow + " | count:  " + count;
                textBox21.Text = datetimenow.ToString();
            }
        }
        public void DeleteTable(string table,string strConnect)
        {
            string sql_delete = "delete from "+ table;

            SqlConnection con = new SqlConnection(strConnect);
            SqlCommand cmd = new SqlCommand(sql_delete, con);
            try
            {
                con.Open();
                cmd.ExecuteNonQuery();
            }
            catch (SqlException ex)
            {
            }
            finally
            {
                con.Close();
            }
        }
        public void scannerVW_GW_SERVER()
        {
            //scandate 15/09/2020
            using (OracleConnection con = new OracleConnection(strdbisc))
            {
                try
                {
                    DeleteTable("VW_GW_SERVER", strSql5);
                    con.Open();
                    int count = 0;
                    string sql = "select * from VW_GW_SERVER ";
                    OracleCommand cmd = new OracleCommand(sql, con);

                    OracleDataReader dr = cmd.ExecuteReader();

                    if (dr.HasRows)
                    {
                        while (dr.Read())
                        {
                            count++;
                            Console.WriteLine("not contains");
                            string scanagaintime = "2020/08/10 13:46:44";
                            string datetimenow2 = DateTime.Now.ToString("yyyy/MM/dd HH:mm:ss");
                            insertDataVW_GW_SERVER(dr[0].ToString(), dr[1].ToString(), dr[2].ToString(), dr[3].ToString(), dr[4].ToString(), dr[5].ToString(), dr[6].ToString(), dr[7].ToString(), datetimenow2.ToString());
                        }
                    }
                    // close and dispose the objects
                    dr.Close();
                    dr.Dispose();
                    cmd.Dispose();
                    con.Close();
                    con.Dispose();

                    string datetimenow = DateTime.Now.ToString("yyyy/MM/dd HH:mm:ss");
                    datetimenow = datetimenow + " | count:  " + count;
                    textBox19.Text = datetimenow.ToString();

                }
                catch (OracleException ex) // catches only Oracle errors
                {
                    MessageBox.Show("The database is unavailable.");
                }
                
            }
        }
        public void scannerVW_GW_PC()
        {
            //scandate 15/09/2020
            using (OracleConnection con = new OracleConnection(strdbisc))
            {
                int count =0;
                try
                {
                    DeleteTable("VW_GW_PC", strSql5);
                    con.Open();
                    string sql = "select * from VW_GW_PC";
                    OracleCommand cmd = new OracleCommand(sql, con);

                    OracleDataReader dr = cmd.ExecuteReader();

                    if (dr.HasRows)
                    {
                        while (dr.Read())
                        {
                            count++;
                            Console.WriteLine("not contains");
                            string scanagaintime = "2020/08/10 13:46:44";
                            string datetimenow2 = DateTime.Now.ToString("yyyy/MM/dd HH:mm:ss");
                            insertDataVW_GW_PC(dr[0].ToString(), dr[1].ToString(), dr[2].ToString(), dr[3].ToString(), dr[4].ToString(), dr[5].ToString(), dr[6].ToString(), dr[7].ToString(), datetimenow2.ToString());
                        }
                    }
                    // close and dispose the objects
                    dr.Close();
                    dr.Dispose();
                    cmd.Dispose();
                    con.Close();
                    con.Dispose();

                }
                catch (OracleException ex) // catches only Oracle errors
                {
                    MessageBox.Show("The database is unavailable.");
                }
                string datetimenow = DateTime.Now.ToString("yyyy/MM/dd HH:mm:ss");
                datetimenow = datetimenow + " | count: " + count;
                textBox20.Text = datetimenow.ToString();
            }
        }
        //canner 14h
        public void scannerCHECK_OVERTIME_DUTY()
        {
            using (SqlConnection con = new SqlConnection(strSql1))
            {
               
                try
                {               
                    con.Open();
                    int count = 0;
                    string sql = "select EMP_NO,NAME,DEPT,DUTYDATE,ISOVERTIME,F_DUTYDATE,ACTUALOVERTIMECOUNT,OVERTIME_S,OVERTIME_E,duty from VW_OVERTIME_DUTY";
                    SqlCommand cmd = new SqlCommand(sql, con);

                    SqlDataReader dr = cmd.ExecuteReader();
                    if (dr.HasRows)
                    {
                        
                        while (dr.Read())
                        {
                            count++;
                            string scanagaintime = "2020-12-21 14:00:18.000";
                            string datetimenow2 = DateTime.Now.ToString("yyyy/MM/dd HH:mm:ss");
                            insertDataCHECK_OVERTIME_DUTY(dr[0].ToString(), dr[1].ToString(), dr[2].ToString(), dr[3].ToString(), dr[4].ToString(), dr[5].ToString(), dr[6].ToString(), dr[7].ToString(), dr[8].ToString(), dr[9].ToString(), scanagaintime, "CHECK_OVERTIME_DUTY");
                        }
                    }
                    // close and dispose the objects
                    dr.Close();
                    dr.Dispose();
                    cmd.Dispose();
                    con.Close();
                    con.Dispose();

                    string datetimenow = DateTime.Now.ToString("yyyy/MM/dd HH:mm:ss");
                    datetimenow = datetimenow + " | count:  " + count;
                    textBox25.Text = datetimenow.ToString();

                }
                catch (OracleException ex) // catches only Oracle errors
                {
                    MessageBox.Show("The database is unavailable."+ex.Message);
                }

                

            }
        }
        public void scannerSFIS_SECSSION_GOLD()
        {
            using (OracleConnection con = new OracleConnection(strToad7))
            {
                 
                try
                {
                    con.Open();
                    int count = 0;
                    // DeleteTable("SFIS_SECSSION_GOLD", strCn);
                    string sql = "select count (*) AS COUNT,  'sfcapi' AS TYPE from v$session";
                    OracleCommand cmd = new OracleCommand(sql, con);
                    OracleDataReader dr = cmd.ExecuteReader();
                    if (dr.HasRows)
                    {
                        while (dr.Read())
                        {
                            count++;
                            string scanagaintime = "2020/08/10 13:46:44";
                            string datetimenow2 = DateTime.Now.ToString("yyyy/MM/dd HH:mm:ss");
                            insertSFIS_SECSSION_GOLD(dr[0].ToString(), dr[1].ToString(), datetimenow2.ToString());
                        }
                    }
                    // close and dispose the objects
                    dr.Close();
                    dr.Dispose();
                    cmd.Dispose();
                    con.Close();
                    con.Dispose();

                    string datetimenow = DateTime.Now.ToString("yyyy/MM/dd HH:mm:ss");
                    datetimenow = datetimenow + " | count:  " + count;
                    textBox26.Text = datetimenow.ToString();

                }
                catch (OracleException ex) // catches only Oracle errors
                {
                    MessageBox.Show("The database is unavailable.");
                }
            }
        }
        public void insertSFIS_SECSSION_GOLD(string COUNT, string TYPE, string SCAN_TIME)
        {

            string sql_send = "Insert SFIS_SECSSION_GOLD ";
            sql_send += " (COUNT,TYPE,SCAN_TIME)";
            sql_send += " Values ('" + COUNT + "', '" + TYPE + "','" + SCAN_TIME + "')";

            SqlConnection con = new SqlConnection(strCn);
            SqlCommand cmd = new SqlCommand(sql_send, con);
            try
            {
                con.Open();
                cmd.ExecuteNonQuery();
            }
            catch (SqlException ex)
            {
            }
            finally
            {
                con.Close();
            }

        }
        public void insertDataCHECK_OVERTIME_DUTY(string EMP_NO, string NAME, string DEPT, string DUTYDATE, string ISOVERTIME, string F_DUTYDATE, string ACTUALOVERTIMECOUNT, string OVERTIME_S, string OVERTIME_E, string duty,string SCAN_DATE,string Table)
        {

            string sql_send = "Insert into " + Table;
            sql_send += " (EMP_NO,NAME,DEPT,DUTYDATE,ISOVERTIME,F_DUTYDATE,ACTUALOVERTIMECOUNT,OVERTIME_S,OVERTIME_E,duty,SCAN_DATE)";
            sql_send += " Values ('" + EMP_NO + "', N'" + NAME + "',N'" + DEPT + "', '" + DUTYDATE + "', '" + ISOVERTIME + "', '" + F_DUTYDATE + "','" + ACTUALOVERTIMECOUNT + "','" + OVERTIME_S + "','"+ OVERTIME_E + "','"+duty+"', '" + SCAN_DATE + "')";

            using (SqlConnection con = new SqlConnection(strSql1))
            {
                SqlCommand cmd = new SqlCommand(sql_send, con);
                try
                {
                    con.Open();
                    cmd.ExecuteNonQuery();
                }
                catch (SqlException ex)
                {
                }
                finally
                {
                    con.Close();
                }
            }
               

        }
        public void insertDataVW_CHROME_VERSION(string PCNAME, string USERNAME, string CUST_KP_NO, string EMP_NO, string INPUT_DATE, string IPS, string HR, string FACTORY, string SCAN_DATE,string Table)
        {

            string sql_send = "Insert into "+Table;
            sql_send += " (PCNAME,USERNAME,CUST_KP_NO,EMP_NO,INPUT_DATE,IPS,HR,FACTORY,SCAN_DATE)";
            sql_send += " Values ('" + PCNAME + "', '" + USERNAME + "','" + CUST_KP_NO + "', '" + EMP_NO + "', '" + INPUT_DATE + "', '" + IPS + "','" + HR + "','" + FACTORY + "', '" + SCAN_DATE + "')";

            SqlConnection con = new SqlConnection(strSql5);
            SqlCommand cmd = new SqlCommand(sql_send, con);
            try
            {
                con.Open();
                cmd.ExecuteNonQuery();
            }
            catch (SqlException ex)
            {
            }
            finally
            {
                con.Close();
            }

        }
        public void insertDataVW_GW_PC(string PCNAME, string USERNAME, string CUST_KP_NO, string EMP_NO, string INPUT_DATE, string IPS, string HR, string FACTORY, string SCAN_DATE)
        {

            string sql_send = "Insert into dbo.VW_GW_PC";
            sql_send += " (PCNAME,USERNAME,CUST_KP_NO,EMP_NO,INPUT_DATE,IPS,HR,FACTORY,SCAN_DATE)";
            sql_send += " Values ('" + PCNAME + "', '" + USERNAME + "','" + CUST_KP_NO + "', '" + EMP_NO + "', '" + INPUT_DATE + "', '" + IPS + "','" + HR + "','" + FACTORY + "', '" + SCAN_DATE + "')";

            SqlConnection con = new SqlConnection(strSql5);
            SqlCommand cmd = new SqlCommand(sql_send, con);
            try
            {
                con.Open();
                cmd.ExecuteNonQuery();
            }
            catch (SqlException ex)
            {
            }
            finally
            {
                con.Close();
            }

        }
        public void insertDataVW_GW_SERVER(string PCNAME,string USERNAME,string CUST_KP_NO,string EMP_NO,string INPUT_DATE,string IPS,string HR,string FACTORY,string SCAN_DATE)
        {

            string sql_send = "Insert into dbo.VW_GW_SERVER";
            sql_send += " (PCNAME,USERNAME,CUST_KP_NO,EMP_NO,INPUT_DATE,IPS,HR,FACTORY,SCAN_DATE)";
            sql_send += " Values ('" + PCNAME + "', '" + USERNAME + "','" + CUST_KP_NO + "', '" + EMP_NO + "', '" + INPUT_DATE + "', '" + IPS + "','" + HR + "','" + FACTORY + "', '" + SCAN_DATE + "')";

            SqlConnection con = new SqlConnection(strSql5);
            SqlCommand cmd = new SqlCommand(sql_send, con);
            try
            {
                con.Open();
                cmd.ExecuteNonQuery();
            }
            catch (SqlException ex)
            {
            }
            finally
            {
                con.Close();
            }

        }
        public void insertDataMysql(string Emp_no,string Emp_name,string depart)
        {
            string sql_send = "Insert into dbo.emp_esd";
            sql_send += " (Emp_no,Emp_name,depart)";
            sql_send += " Values ('" + Emp_no + "', N'" + Emp_name + "',N'" + depart +"')";
            SqlConnection con = new SqlConnection(strSql4);
            SqlCommand cmd = new SqlCommand(sql_send, con);
            try
            {
                con.Open();
                cmd.ExecuteNonQuery();
            }
            catch (SqlException ex)
            {
            }
            finally
            {
                con.Close();
            }
        }

        public void insertDataMysql_emp_esd_all(string f_empno, string f_plantno, string f_name, string f_site, string f_organno, string f_department, string f_intogroup, string f_grand, string f_manage, string f_emptype_use, string f_flid, string f_istaiwan, string f_telephone, string f_ext, string line_leader, string utc_f_untogroup)
        {

            string sql_send = "Insert into dbo.emp_esd_all";
            sql_send += " (f_empno,f_plantno,f_name,f_site,f_organno,f_department,f_intogroup,f_grand,f_manage,f_emptype_use,f_flid,f_istaiwan,f_telephone,f_ext,line_leader,utc_f_untogroup)";
            sql_send += " Values ('" + f_empno + "','" + f_plantno + "',N'" + f_name + "',  N'" + f_site + "','" + f_organno + "',N'" + f_department + "','" + f_intogroup + "',N'" + f_grand + "',N'" + f_manage + "','" + f_emptype_use + "','" + f_flid + "','" + f_istaiwan + "','" + f_telephone + "','" + f_ext + "','" + line_leader + "','"+ utc_f_untogroup + "')";

            SqlConnection con = new SqlConnection(strSql4);
            SqlCommand cmd = new SqlCommand(sql_send, con);
            try
            {
                con.Open();
                cmd.ExecuteNonQuery();
            }
            catch (SqlException ex)
            {
            }
            finally
            {
                con.Close();
            }

        }
        //insert pc 15/09/2020
        //insert NTP 10/26/2020
        public void insertDataNTP(string PCNAME, string USERNAME, string NTP, string NOW, string RANGETIME, string EMP_NO,string CITY_LOMO,string MM ,string CUST_KP_NO,string INPUT_dATE, string SCAN_DATE)
        {

            string sql_send = "Insert into NTP_CLIENT_TIME ";
            sql_send += " (PCNAME,USERNAME,NTP,NOW,RANGETIME,EMP_NO,CITY_LOMO,MM,CUST_KP_NO,INPUT_dATE,SCAN_DATE)";
            sql_send += " Values ('" + PCNAME + "', '" + USERNAME + "','" + NTP + "','" + NOW + "','" + RANGETIME + "','" + EMP_NO + "','"+ CITY_LOMO + "','"+MM+"','"+ CUST_KP_NO + "','"+ INPUT_dATE + "','" + SCAN_DATE + "')";

            SqlConnection con = new SqlConnection(strCn);
            SqlCommand cmd = new SqlCommand(sql_send, con);
            try
            {
                con.Open();
                cmd.ExecuteNonQuery();
            }
            catch (SqlException ex)
            {
            }
            finally
            {
                con.Close();
            }

        }
        //insert pc 15/09/2020
        public void insertDataPcNotOpen(string MACID, string IP, string RANGEIP, string EMP_NO, string INPUT_DATE,string PCNAME,string HR,string LOCATION, string SCANDATE)
        {

            string sql_send = "Insert into C_PC_NOTOPEN";
            sql_send += " (MACID,IP,RANGEIP,EMP_NO,INPUT_DATE,PCNAME,HR,LOCATION,SCAN_DATE)";
            sql_send += " Values ('" + MACID + "', '" + IP + "','" + RANGEIP + "','" + EMP_NO + "','" + INPUT_DATE + "','"+PCNAME+"','"+ HR + "','"+ LOCATION + "','" + SCANDATE + "')";

            SqlConnection con = new SqlConnection(strCn);
            SqlCommand cmd = new SqlCommand(sql_send, con);
            try
            {
                con.Open();
                cmd.ExecuteNonQuery();
            }
            catch (SqlException ex)
            {
            }
            finally
            {
                con.Close();
            }

        }
        //INSERT PC_CONTROL
        public void insertDataPcControl(string PCNAME, string USERNAME, string IPS, string HR, string INPUT_DATE, string SCANDATE)
        {

            string sql_send = "Insert into C_PC_CONTROL_T ";
            sql_send += " (PCNAME,USERNAME,IPS,HR,INPUT_DATE,SCANDATE)";
            sql_send += " Values ('" + PCNAME + "', '" + USERNAME + "','" + IPS + "','" + HR + "','" + INPUT_DATE + "','" + SCANDATE + "')";

            SqlConnection con = new SqlConnection(strCn);
            SqlCommand cmd = new SqlCommand(sql_send, con);
            try
            {
                con.Open();
                cmd.ExecuteNonQuery();
            }
            catch (SqlException ex)
            {
            }
            finally
            {
                con.Close();
            }

        }

        //insert sql chi than 
        public void insertDataDailyMail(string MAIL_ID, string LOCATION, string SERVER_IP, string DATABASE_NAME, string PROGRAM_NAME, string DESCRIBE, string OWNER, string TIME, string EMP_NAME,string SCAN_DATE)
        {

            string sql_send = "Insert into CHECK_DAILY ";
            sql_send += " (MAIL_ID,LOCATION,SERVER_IP,DATABASE_NAME,PROGRAM_NAME,DESCRIBE,OWNER,TIME,EMP_NAME,SCAN_DATE)";
            sql_send += " Values ('" + MAIL_ID + "', '" + LOCATION + "','" + SERVER_IP + "','" + DATABASE_NAME + "','" + PROGRAM_NAME + "','" + DESCRIBE + "','" + OWNER + "','" + TIME + "',N'"+EMP_NAME+"','"+ SCAN_DATE + "')";

            SqlConnection con = new SqlConnection(strCn);
            SqlCommand cmd = new SqlCommand(sql_send, con);
            try
            {
                con.Open();
                cmd.ExecuteNonQuery();
            }
            catch (SqlException ex)
            {
            }
            finally
            {
                con.Close();
            }

        }
        //INSERT sql chi than 19/08/2020
        public void insertDataCheckEmp(string EMP_NO, string EMP_NAME, string ID, string AMOUNT, string CHECK_FROM, string CHECK_TO, string SCAN_DATE)
        {

            string sql_send = "Insert into CHECK_EMP ";
            sql_send += " (EMP_NO,EMP_NAME,ID,AMOUNT,CHECK_FROM,CHECK_TO,SCAN_DATE)";
            sql_send += " Values ('" + EMP_NO + "', N'" + EMP_NAME + "','" + ID + "','" + AMOUNT + "','" + CHECK_FROM + "','" + CHECK_TO + "','" + SCAN_DATE + "')";

            SqlConnection con = new SqlConnection(strCn);
            SqlCommand cmd = new SqlCommand(sql_send, con);
            try
            {
                con.Open();
                cmd.ExecuteNonQuery();
            }
            catch (SqlException ex)
            {
            }
            finally
            {
                con.Close();
            }

        }
        public void insertdatasql(string pcname, string username, string cust_kp_no, string emp_no, string input_date,string ips, string hr, string factory, string scan_date, string nametable)
        {
            
            string sql_send = "Insert into " +nametable;
            sql_send += " (PCNAME,USERNAME,CUST_KP_NO,EMP_NO,INPUT_DATE,IPS,HR,FACTORY,SCAN_DATE)";
            sql_send += " Values ('" + pcname + "', '" + username + "','" + cust_kp_no + "','" + emp_no + "','','" + ips + "','" + hr + "','" + factory + "','" + scan_date + "')";

            SqlConnection con = new SqlConnection(strCn);
            SqlCommand cmd = new SqlCommand(sql_send, con);
            try
            {
                con.Open();
                cmd.ExecuteNonQuery();
            }
            catch (SqlException ex)
            {
            }
            finally
            {
                con.Close();
            }
                   
        }
        public void insertdatadetectportsql(string pcname, string username, string cust_kp_no, string emp_no, string input_date, string ips, string hr, string factory, string scan_date, string nametable, string docno, string expiration, string ck)
        {

            string sql_send = "Insert into " + nametable;
            sql_send += " (PCNAME,USERNAME,CUST_KP_NO,EMP_NO,INPUT_DATE,IPS,HR,FACTORY,SCAN_DATE,DOC_NO,EXPIRATION_DATE,CK)";
            sql_send += " Values ('" + pcname + "', '" + username + "','" + cust_kp_no + "','" + emp_no + "','','" + ips + "','" + hr + "','" + factory + "','" + scan_date + "','" + docno + "','" + expiration + "','" + ck + "')";

            SqlConnection con = new SqlConnection(strCn);
            SqlCommand cmd = new SqlCommand(sql_send, con);
            try
            {
                con.Open();
                cmd.ExecuteNonQuery();
            }
            catch (SqlException ex)
            {
            }
            finally
            {
                con.Close();
            }
        }
        public void insertdataTREND_LINE_PACKAGE(string window, int count, string input_date, string table)
        {

            string sql_send = "Insert into "+table;
            sql_send += " (WINDOW,COUNT,SCAN_DATE)";
            sql_send += " Values ('" + window + "', " + count + "  ,'" + input_date + "')";

            SqlConnection con = new SqlConnection(strCn);
            SqlCommand cmd = new SqlCommand(sql_send, con);
            try
            {
                con.Open();
                cmd.ExecuteNonQuery();
            }
            catch (SqlException ex)
            {
            }
            finally
            {
                con.Close();
            }
        }
        public void  insertdataTREND_LINE_PACKAGE_DETAIL(string CUST_KP_NO, string EMP_NO, string CITY_LOMO, string CITY_WORD, string PC_NAME, string USER_NAME, string WINDOW,string IPS,  string SCAN_DATE,string table)
        {

            string sql_send = "Insert into "+table;
            sql_send += " (CUST_KP_NO,EMP_NO,CITY_LOMO,CITY_WORD,PCNAME,USERNAME,WINDOW,IPS,SCAN_DATE)";
            sql_send += " Values ('" + CUST_KP_NO + "', '" + EMP_NO + "','" + CITY_LOMO + "',N'"+ CITY_WORD + "','" + PC_NAME + "','" + USER_NAME + "','" + WINDOW + "','" + IPS + "','" + SCAN_DATE + "')";

            SqlConnection con = new SqlConnection(strCn);
            SqlCommand cmd = new SqlCommand(sql_send, con);
            try
            {
                con.Open();
                cmd.ExecuteNonQuery();
            }
            catch (SqlException ex)
            {
            }
            finally
            {
                con.Close();
            }
        }
        public void insertdataVW_ROOT_PACKAGE(string PCNAME,string IPS,string HR,string USERNAME, string WINDOW,string SCAN_DATE)
        {
            string sql_send = "Insert into VW_ROOT_PACKAGE";
            sql_send += " (PCNAME,IPS,HR,USERNAME,WINDOW,SCAN_DATE)";
            sql_send += " Values ('" + PCNAME + "', '" + IPS + "','" + HR + "',N'" + USERNAME + "','" + WINDOW + "','" + SCAN_DATE + "')";

            SqlConnection con = new SqlConnection(strCn);
            SqlCommand cmd = new SqlCommand(sql_send, con);
            try
            {
                con.Open();
                cmd.ExecuteNonQuery();
            }
            catch (SqlException ex)
            {
            }
            finally
            {
                con.Close();
            }
        }
        public void insertdatasqlupdatewindows(string macaddress, string ipaddress, string cust_kp_no, string emp_no, string input_date, string os,string factory, string scan_date, string nametable)
        {

            string sql_send = "Insert into " + nametable;
            sql_send += " (MACADDRESS,IPADDRESS,CUST_KP_NO,EMP_NO,INPUT_DATE,OS,FACTORY,SCAN_DATE)";
            sql_send += " Values ('" + macaddress + "', '" + ipaddress + "','" + cust_kp_no + "','" + emp_no + "','','" + os + "','" + factory + "','" + scan_date + "')";

            SqlConnection con = new SqlConnection(strCn);
            SqlCommand cmd = new SqlCommand(sql_send, con);
            try
            {
                con.Open();
                cmd.ExecuteNonQuery();
                Console.WriteLine("Records Inserted Successfully");
            }
            catch (SqlException ex)
            {
                Console.WriteLine("Error Generated. Details: " + ex.ToString());
            }
            finally
            {
                con.Close();
            }

        }
        //Run with window
        private void RegisterInStartup()
        {
            RegistryKey registryKey = Registry.CurrentUser.OpenSubKey
                    ("SOFTWARE\\Microsoft\\Windows\\CurrentVersion\\Run", true);
            registryKey.SetValue("SaveReportData", Application.ExecutablePath);
        }
        public void deletedata(string nametable)
        {
            DialogResult result = MessageBox.Show("Do you want to reset database ?", "Reset database !  ", MessageBoxButtons.OKCancel, MessageBoxIcon.Information);
            if (result.Equals(DialogResult.OK))
            {

                string sql_send = "DELETE FROM " + nametable + " where SCAN_DATE > '2020-07-24 00:00' and SCAN_DATE < '2020-07-24 23:00' ";
                SqlConnection con = new SqlConnection(strCn);
                SqlCommand cmd = new SqlCommand(sql_send, con);
                try
                {
                    con.Open();
                    cmd.ExecuteNonQuery();
                    Console.WriteLine("SMOAMLERTDB Successfully");
                    MessageBox.Show("Delete Successfully table : " + nametable);
                }
                catch (SqlException ex)
                {
                    Console.WriteLine("Error Generated. Details: " + ex.ToString());
                    MessageBox.Show("Delete not  Successfully : " + nametable);
                }
                finally
                {
                    con.Close();
                }
            }
            else
            {
            }
        }

        [Obsolete]
        public void scannerdaily()
        {
            scannerdata("VW_NO_INSTALL_SYMANTEC");
            scannerdata("VW_JOIN_DOMAIN");
            scannerdata("VW_UPDATE_VERSION_SYMANTEC");
            scannerdata("VW_OPEN_PORT_DOCNO");
            scannerdata("VW_USB_AVAI_UNAVAI");
            scannerdata("VW_USE_PROXY");
            scannerdata("VW_USE_CD");
            scannerdata("VW_CONNECT_INTERNET");
            scannerUpdateWindowStatusdata("VW_UPDATE_WINDOW_STATUS_ALL");
            scannerdata("VW_DEFINE_SOFWARE");
            scannerDataNTP(); 
            scannerdataTREND_LINE_PACKAGE_DETAIL("SELECT  cust_kp_no,EMP_NO,city_lomo,city_word,pcname,username,window,ips  FROM TREND_LINE_PACKAGE_DETAIL", "TREND_LINE_PACKAGE_DETAIL");
            scannerdataVW_ROOT_PACKAGE();
            scannerdataTREND_LINE_PACKAGE("select count(distinct pcname) as counts, 'Windows 10' as window , getdate() as times from [dbo].[TREND_LINE_PACKAGE_DETAIL] where window ='Windows 10' and convert(date, SCAN_DATE, 110) = convert(date, getdate(), 110) union all select count(distinct pcname) as counts,'Server 2012' as window, getdate() as times from[dbo].[TREND_LINE_PACKAGE_DETAIL] where window = 'Server 2012' and convert(date, SCAN_DATE, 110) = convert(date, getdate(), 110) union all select count(distinct pcname) as counts, 'Server 2016' as window , getdate() as times from[dbo].[TREND_LINE_PACKAGE_DETAIL] where window = 'Server 2016' and convert(date, SCAN_DATE, 110) = convert(date, getdate(), 110) union all select count(distinct pcname) as counts,'Server 2019' as window ,getdate() as times  from[dbo].[TREND_LINE_PACKAGE_DETAIL] where window = 'Server 2019' and convert(date, SCAN_DATE, 110) = convert(date, getdate(), 110)", "TREND_LINE_PACKAGE");
            scannerdataTREND_LINE_PACKAGE_DETAIL("select cust_kp_no,EMP_NO,city_lomo,city_word,pcname,username,window,ips from GZF_TREND_LINE_PACKAGE_DETAIL", "GZF_TREND_LINE_PACKAGE_DETAIL");
            scannerdataTREND_LINE_PACKAGE("select  count(distinct pcname) as counts, 'Windows 10' as window , getdate() as times from [dbo].[GZF_TREND_LINE_PACKAGE_DETAIL] where window ='Windows 10' and  convert(date,SCAN_DATE,110) = convert(date,getdate(),110) union all select   count(distinct pcname) as counts,'Server 2012' as window, getdate() as times from [dbo].[GZF_TREND_LINE_PACKAGE_DETAIL] where window ='Server 2012' and  convert(date,SCAN_DATE,110) = convert(date,getdate(),110) union all select   count(distinct pcname) as counts, 'Server 2016' as window , getdate() as times from [dbo].[GZF_TREND_LINE_PACKAGE_DETAIL] where window ='Server 2016' and  convert(date,SCAN_DATE,110) = convert(date,getdate(),110) union all select   count(distinct pcname) as counts,'Server 2019' as window ,getdate() as times  from [dbo].[GZF_TREND_LINE_PACKAGE_DETAIL] where window ='Server 2019' and convert(date,SCAN_DATE,110) = convert(date,getdate(),110)", "GZF_TREND_LINE_PACKAGE");
            scannerdata("vw_server_capacity");
            scannerdata("VW_NUMBER_CARD");
            scannerdata("VW_DEFINE_BLUETOOL");
            scannerdata("VW_KERNEL_VERSION");
        }
        //auto scanner data than
        public void scannerdaily1()
        {
            scannerCHECK_DAILY();
            scannerCHECK_EMP();
            scannerCHECK_EMP1();
            scannerCHECK_EMP2();
            scannerCHECK_EMP3();
            //scannerCHECK_EMP4();
            scannerCHECK_EMP5();
            scannerCHECK_EMP6();
            scannerCHECK_EMP7();
            scannerCHECK_EMP8();
            scannerC_PC_CONTROL_T();
            scannerC_PC_NOTOPEN();
            scannerVW_GW_SERVER();
            scannerVW_GW_PC();
            scannerVW_CHROME_VERSION();
            scannerVW_GZ_PC();
            scannerVW_HT_PC();
            //scannerMysql_emp_esd();
            //scannerMysql_emp_esd_all();
        }

        [Obsolete]
        public void rundaily()
        {
            issend = true;
            lineChanger("true", 1);
            scannerdaily();
            Console.WriteLine("check daily  ==============================================================!");
            
        }
        public void rundaily1()
        {
            issend2 = true;
            lineChanger1("true", 1);
            scannerdaily1();
            Console.WriteLine("check daily  ==============================================================!");
            
        }
        public void rundaily14h()
        {
            issend3 = true;
            lineChanger1("true", 2);
            scannerCHECK_OVERTIME_DUTY();
            Console.WriteLine("check daily  ==============================================================!");
        }
        public void rundaily9h()
        {
            issend4 = true;
            lineChanger1("true", 3);
            scannerSFIS_SECSSION_GOLD();
            Console.WriteLine("check daily  ==============================================================!");

        }
        static void lineChanger(string newText, int line_to_edit)
        {
            string fileLPath = AppDomain.CurrentDomain.BaseDirectory + "\\saveinfor.txt";
            string[] arrLine = File.ReadAllLines(fileLPath);
            arrLine[line_to_edit - 1] = newText;
            System.IO.File.WriteAllLines(fileLPath, arrLine);
        }
        static void lineChanger1(string newText, int line_to_edit)
        {
            string fileLPath = AppDomain.CurrentDomain.BaseDirectory + "\\saveinfor.txt";
            string[] arrLine = File.ReadAllLines(fileLPath);
            arrLine[line_to_edit] = newText;
            System.IO.File.WriteAllLines(fileLPath, arrLine);
        }

        public void getinforexception()
        {
            string filePath2 = AppDomain.CurrentDomain.BaseDirectory + "\\saveexception2.txt"; //1

            string[] lines2;

            if (System.IO.File.Exists(filePath2))
            {
                lines2 = System.IO.File.ReadAllLines(filePath2);
                listexception = lines2[0].ToString();
            }
            else
            {

            }
        }
        private void addListIpExceptionToolStripMenuItem_Click(object sender, EventArgs e)
        {
            ExceptionIp exception = new ExceptionIp();
            exception.ShowDialog();
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            getinforexception();
        }

        [Obsolete]
        private void scandaily_Tick(object sender, EventArgs e)
        {
            Console.WriteLine("check send infor daily  !!!!!!!");
            //Time when method needs to be called 4:00h
            var DailyTime = "04:00:00";
            var timeParts = DailyTime.Split(new char[1] { ':' });

            var dateNow = DateTime.Now;
            var date = new DateTime(dateNow.Year, dateNow.Month, dateNow.Day,
                       int.Parse(timeParts[0]), int.Parse(timeParts[1]), int.Parse(timeParts[2]));
            Console.WriteLine("Time now is:" + dateNow.ToString());
            if (dateNow > date)
            {
                Console.WriteLine("====================== dateNow > date ============");
                Console.WriteLine("issend is :" + issend.ToString());
                if (issend == false)
                {
                    //auto run
                    rundaily();
                    label7.Text = "Run scanner daily  at :" + dateNow.ToString();
                }
                label8.Text = "Check scanner daily at : " + dateNow.ToString();
            }
            else
            {
                issend = false;
                Console.WriteLine("issend is :" + issend.ToString());
                lineChanger("false", 1);
                Console.WriteLine("dont't send ");
                label8.Text = "Check scanner daily at :" + dateNow.ToString();
            }
        }
        //run daily 7h
        private void timer1_Tick(object sender, EventArgs e)
        {
            Console.WriteLine("check send infor daily  !!!!!!!");
            //Time when method needs to be called 7:30h

            var DailyTime = "07:30:00";
            var timeParts = DailyTime.Split(new char[1] { ':' });

            var dateNow = DateTime.Now;
            var date = new DateTime(dateNow.Year, dateNow.Month, dateNow.Day,
                       int.Parse(timeParts[0]), int.Parse(timeParts[1]), int.Parse(timeParts[2]));
            Console.WriteLine("Time now is:" + dateNow.ToString());
            if (dateNow > date)
            {
                Console.WriteLine("====================== dateNow > date ============");
                Console.WriteLine("issend is :" + issend2.ToString());
                if (issend2 == false)
                {
                    //auto run;
                    rundaily1();
                    label7.Text = "Run scanner daily  at :" + dateNow.ToString();
                }
                label8.Text = "Check scanner daily at : " + dateNow.ToString();
            }
            else
            {
                issend2 = false;
                Console.WriteLine("issend is :" + issend2.ToString());
                lineChanger("false", 2);
                Console.WriteLine("dont't send ");
                label8.Text = "Check scanner daily at :" + dateNow.ToString();
            }
        }
        //run daily 14h
        private void Daily14h_Tick(object sender, EventArgs e)
        {
            Console.WriteLine("check send infor daily  !!!!!!!");
            //Time when method needs to be called
            var DailyTime = "14:00:00";
            var timeParts = DailyTime.Split(new char[1] { ':' });

            var dateNow = DateTime.Now;
            var date = new DateTime(dateNow.Year, dateNow.Month, dateNow.Day,
                       int.Parse(timeParts[0]), int.Parse(timeParts[1]), int.Parse(timeParts[2]));
            Console.WriteLine("Time now is:" + dateNow.ToString());
            if (dateNow > date)
            {
                Console.WriteLine("====================== dateNow > date ============");
                Console.WriteLine("issend is :" + issend3.ToString());
                if (issend3 == false)
                {
                    rundaily14h();
                    label7.Text = "Run scanner daily  at :" + dateNow.ToString();
                }
                label8.Text = "Check scanner daily at : " + dateNow.ToString();
            }
            else
            {
                issend3 = false;
                Console.WriteLine("issend is :" + issend3.ToString());
                lineChanger("false", 3);
                Console.WriteLine("dont't send ");
                label8.Text = "Check scanner daily at :" + dateNow.ToString();
            }
        }
        //scan 2h once
        private void Scan2once_Tick(object sender, EventArgs e)
        {
            Console.WriteLine("check send infor daily  !!!!!!!");
            //Time when method needs to be called
            var DailyTime = "09:00:00";
            var timeParts = DailyTime.Split(new char[1] { ':' });

            var dateNow = DateTime.Now;
            var date = new DateTime(dateNow.Year, dateNow.Month, dateNow.Day,
                       int.Parse(timeParts[0]), int.Parse(timeParts[1]), int.Parse(timeParts[2]));
            Console.WriteLine("Time now is:" + dateNow.ToString());
            if (dateNow > date)
            {
                Console.WriteLine("====================== dateNow > date ============");
                Console.WriteLine("issend is :" + issend4.ToString());
                if (issend4 == false)
                {
                    rundaily9h();
                    label7.Text = "Run scanner daily  at :" + dateNow.ToString();
                }
                label8.Text = "Check scanner daily at : " + dateNow.ToString();
            }
            else
            {
                issend4 = false;
                Console.WriteLine("issend is :" + issend4.ToString());
                lineChanger("false", 4);
                Console.WriteLine("dont't send ");
                label8.Text = "Check scanner daily at :" + dateNow.ToString();
            }
            
        }

        [Obsolete]
        private void button1_Click(object sender, EventArgs e)
        {
            scannerdata("VW_NUMBER_CARD");
            scannerdata("VW_DEFINE_BLUETOOL");
            scannerdata("VW_KERNEL_VERSION");
            //rundaily();
            //MessageBox.Show("insert ok");
            //rundaily1();
            //MessageBox.Show("insert 2");
            //scannerdata("VW_DEFINE_SOFWARE");

            string where = "SCAN_DATE > '2021-02-17 00:00:00.000'";

            //DeleteTableWhere("VW_NO_INSTALL_SYMANTEC", strCn, where);
            //DeleteTableWhere("VW_JOIN_DOMAIN", strCn, where);
            //DeleteTableWhere("VW_UPDATE_VERSION_SYMANTEC", strCn, where);
            //DeleteTableWhere("VW_OPEN_PORT_DOCNO", strCn, where);
            //DeleteTableWhere("VW_USB_AVAI_UNAVAI", strCn, where);
            //DeleteTableWhere("VW_USE_PROXY", strCn, where);
            //DeleteTableWhere("VW_USE_CD", strCn, where);
            //DeleteTableWhere("VW_CONNECT_INTERNET", strCn, where);
            //DeleteTableWhere("VW_UPDATE_WINDOW_STATUS_ALL", strCn, where);
            //DeleteTableWhere("VW_DEFINE_SOFWARE", strCn, where);
            //DeleteTableWhere("NTP_CLIENT_TIME", strCn, where);
            //DeleteTableWhere("TREND_LINE_PACKAGE_DETAIL", strCn, where);
            //DeleteTableWhere("VW_ROOT_PACKAGE", strCn, where);
            //DeleteTableWhere("TREND_LINE_PACKAGE", strCn, where);
            //DeleteTableWhere("vw_server_capacity", strCn, where);
            //DeleteTableWhere("VW_NUMBER_CARD", strCn, where);
            //DeleteTableWhere("VW_DEFINE_BLUETOOL", strCn, where);
            //DeleteTableWhere("VW_KERNEL_VERSION", strCn, where);
            //MessageBox.Show("Xoa 1 ok");

            ////////rundaily1
            //DeleteTableWhere("CHECK_DAILY", strCn, where);
            //DeleteTableWhere("CHECK_EMP", strCn, where);
            //DeleteTableWhere("C_PC_CONTROL_T", strCn, where);
            //DeleteTableWhere("C_PC_NOTOPEN", strCn, where);
            //DeleteTableWhere("VW_GW_SERVER", strSql5, where);
            //DeleteTableWhere("VW_GW_PC", strSql5, where);
            //DeleteTableWhere("VW_CHROME_VERSION", strSql5, where);
            //DeleteTableWhere("VW_GZ_PC", strSql5, where);
            //DeleteTableWhere("VW_HT_PC", strSql5, where);
            //DeleteTableWhere("emp_esd", strSql4, where);
            //MessageBox.Show("Xoa 2 ok");

        }

        public void DeleteTableWhere(string table, string strConnect,string where)
        {
            string sql_delete = "delete from " + table +" where "+where;

            SqlConnection con = new SqlConnection(strConnect);
            SqlCommand cmd = new SqlCommand(sql_delete, con);
            try
            {
                con.Open();
                cmd.ExecuteNonQuery();
            }
            catch (SqlException ex)
            {
            }
            finally
            {
                con.Close();
            }
        }

    }
}
