using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace 个人成绩单生成器
{
    public class Student
    {
        public string id;
        public string className;
        public string sign;
        public string oid;
        public string name;

        //入学时间
        public DateTime dataTime;
        public Student()
        {

        }

        public Student(string id, string className , string sign , string oid, string name, DateTime dataTime)
        {
            this.id = id;
            this.className = className;
            {
                int temp = int.Parse(oid);
                this.oid = String.Format("{0:D8}", temp);
            }
            this.name = name;
            this.sign = sign;
            this.dataTime = dataTime;
        }
    }
}
