using System;
using System.Collections;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Excel;

namespace ExcelDataReader
{
    class Program
    {
        static void Main(string[] args)
        {
            string path = System.Environment.CurrentDirectory;   //获取当前目录
            //在当前目录创建存放csv和txt格式文件的文件夹                                                   
            string csvDir = path + "\\csv";
            string txtDir = path + "\\txt";
            if (!Directory.Exists(csvDir))
            {
                Directory.CreateDirectory(csvDir);
            }
            else
            {
                Directory.Delete(csvDir, true);
                Directory.CreateDirectory(csvDir);
            }
            if (!Directory.Exists(txtDir))
            {
                Directory.CreateDirectory(txtDir);
            }
            else
            {
                Directory.Delete(txtDir, true);
                Directory.CreateDirectory(txtDir);
            }

            //获取当前目录下所有文件，查找到xlsx文件
            DirectoryInfo dir = new DirectoryInfo(path);
            FileInfo[] fileInfos = dir.GetFiles();      //获取当前目录下的所有文件

            //遍历目录下所有excel文件，将其名称存放到excelFileList。将其全路径存放到excelFullPath
            ArrayList excelFileList = new ArrayList();
            ArrayList excelFullPath = new ArrayList();
            for (int i = 0; i < fileInfos.Length; ++i)
            {
                if (fileInfos[i].Name.EndsWith(".xlsx") || fileInfos[i].Name.EndsWith(".xls"))
                {
                    excelFileList.Add(fileInfos[i].ToString());
                    excelFullPath.Add(path + "\\" + fileInfos[i].ToString());
                }
            }
            for (int i = 0; i < excelFullPath.Count; ++i)
            {
                //Console.WriteLine("excelFileList==="+ excelFileList.Count+"|");
                //Console.WriteLine("excelFullPath===" + excelFullPath.Count + "|");
                //Console.WriteLine(excelFullPath[i]);
                FileStream stream = File.Open(excelFullPath[i].ToString(), FileMode.Open, FileAccess.Read);

                IExcelDataReader excelDataReader = excelFullPath[i].ToString().Contains(".xlsx") ? ExcelReaderFactory.CreateOpenXmlReader(stream) : ExcelReaderFactory.CreateBinaryReader(stream);

                DataSet result = excelDataReader.AsDataSet();
                excelDataReader.IsFirstRowAsColumnNames = true;

                //DataTable at = result.Tables[0];
                //Console.WriteLine("DataTable=" + at.Rows[0].ItemArray[0].ToString());

                

                //遍历表格中数据，按行写入csv文件和txt文件
                FileStream temStreamCsv = new FileStream(csvDir + "\\" + excelFileList[i].ToString().Replace(".xlsx", ".csv"), FileMode.OpenOrCreate, FileAccess.ReadWrite);
                FileStream temStreamTxt = new FileStream(txtDir + "\\" + excelFileList[i].ToString().Replace(".xlsx", ".txt"), FileMode.OpenOrCreate, FileAccess.ReadWrite);
                StreamWriter sw = new StreamWriter(temStreamCsv);
                StreamWriter swTxt = new StreamWriter(temStreamTxt, Encoding.UTF8);
                for (int ii = 0; ii < result.Tables[0].Rows.Count; ++ii)
                {
                    string rowStr = String.Empty;
                    for (int j = 0; j < result.Tables[0].Columns.Count; ++j)
                    {
                        if (j != result.Tables[0].Columns.Count - 1)
                        {
                            rowStr += result.Tables[0].Rows[ii][j].ToString() + ",";
                        }
                        else
                        {
                            rowStr += result.Tables[0].Rows[ii][j].ToString();
                        }
                    }
                    sw.WriteLine(rowStr);
                    swTxt.WriteLine(rowStr);
                }
                Console.WriteLine(excelFileList[i]);
                Console.WriteLine("表格行数："+result.Tables[0].Rows.Count);
                Console.WriteLine("转换OK了！！");
                Console.WriteLine("---------------------------------------------------------------------");
                sw.Close();
                swTxt.Close();
                temStreamCsv.Close();
                temStreamTxt.Close();
                excelDataReader.Close();
                stream.Close();
                
            }
            Console.ReadKey();
        }
    }
}
