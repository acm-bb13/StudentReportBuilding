﻿using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using MySql.Data.MySqlClient;

namespace 个人成绩单生成器
{
    class SQLManage
    {
        public static MySqlConnection getConn()
        {
            if (conns == null)
            {
                string connetStr = "server=localhost;port=3306;user=root;password=123456;database=studentmanagesystem;Charset=utf8;";
                //string connetStr = "server=www.bb13.xyz;port=3306;user=root;password=123456;database=studentmanagesystem;Charset=utf8;";
                conns = new MySqlConnection(connetStr);
            }
            return conns;
        }

        public static void closeConn()
        {
            if (conns != null)
            {
                conns.Close();
                conns = null;
            }
        }

        public static MySqlDataReader GetReader(string sql)
        {
            MySqlDataReader mySqlDataReader = null;
            MySqlConnection conn = getConn(); //连接数据库
            conn.Open(); //打开数据库连接
            MySqlCommand cmd = new MySqlCommand(sql, conn);
            mySqlDataReader = cmd.ExecuteReader();
            return mySqlDataReader;
        }


        public static void command(string sql)
        {
            MySqlConnection conn = getConn(); //连接数据库
            conn.Open(); //打开数据库连接
            MySqlCommand cmd = new MySqlCommand(sql, conn);
            cmd.ExecuteNonQuery();
            conn.Close();
            conns = null;
        }

        static MySqlConnection conns = null;
        public static void commands(int com, string sql = "")
        {
            if (com == 1)
            {
                conns = getConn(); //连接数据库
                conns.Open(); //打开数据库连接
            }
            if (com == 2)
            {
                MySqlCommand cmd = new MySqlCommand(sql, conns);
                cmd.ExecuteNonQuery();
            }
            if (com == 3)
            {
                if (conns != null)
                {
                    conns.Close();
                    conns = null;
                }
            }
        }
    }
}
