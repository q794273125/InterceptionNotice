using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.IO;


namespace NVITRC1.Common
{
    public static partial class TxtHelper
    {
        public static void Write(string path, string value)
        {
            if (File.Exists(path))
            {
                File.Delete(path);
            }
            FileStream fs = new FileStream(path, FileMode.Create);
            StreamWriter sw = new StreamWriter(fs, Encoding.UTF8);
            //开始写入
            sw.Write(value);
            //清空缓冲区
            sw.Flush();
            //关闭流
            sw.Close();
            fs.Close();
        }
        public static void Write(string filePath, string fileName, string value)
        {
            if (!Directory.Exists(filePath))
            {
                Directory.CreateDirectory(filePath);
            }
            Write(filePath + "\\" + fileName, value);
        }

        public static string Read(string filePath)
        {
            if (!File.Exists(filePath))
            {
                throw new Exception("路径错误");
            }

            StreamReader reader = new StreamReader(filePath, System.Text.Encoding.UTF8, false);
            string str = reader.ReadToEnd();
            reader.Dispose();
            reader.Close();
            return str;
        }

        public static void WriteAppend(string path, string value)
        {
            if (!File.Exists(path))
            {
                throw new Exception("路径错误");
            }
            FileStream fs = new FileStream(path, FileMode.Append);
            StreamWriter sw = new StreamWriter(fs);
            sw.WriteLine(value);
            sw.Close();
            fs.Close();
        }

        /// <summary>
        /// 用 ASCII 码范围判断字符是不是汉字
        /// </summary>
        /// <param name="text">待判断字符或字符串</param>
        /// <returns>真：是汉字；假：不是</returns>
        public static bool CheckStringChinese(string text)
        {
            bool res = false;
            foreach (char t in text)
            {
                if ((int)t > 127)
                    res = true;
            }
            return res;
        }
    }
}
