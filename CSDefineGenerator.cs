using System;
using System.IO;
using System.Data;
using System.Text;
using System.Collections.Generic;

namespace excel2json
{
    /// <summary>
    /// 根据表头，生成C#类定义数据结构
    /// 表头使用三行定义：字段名称、字段类型、注释
    /// </summary>
    class CSDefineGenerator
    {
        struct FieldDef
        {
            public string name;
            public string type;
            public string value;
            public string comment;

            public bool isArray 
            { 
                get
                {
                    return type.EndsWith("[]");
                } 
            }
        }

        string mCode;

        public string code {
            get {
                return this.mCode;
            }
        }

        public CSDefineGenerator(string excelName, ExcelLoader excel, string excludePrefix)
        {
            //-- 创建代码字符串
            StringBuilder sb = new StringBuilder();
            sb.AppendLine("//");
            sb.AppendLine("// Auto Generated Code By excel2json");
            sb.AppendLine("// https://neil3d.gitee.io/coding/excel2json.html");
            sb.AppendLine("// 1. 每个 Sheet 形成一个 Struct 定义, Sheet 的名称作为 Struct 的名称");
            sb.AppendLine("// 2. 表格约定：第一行是变量名称，第二行是变量类型");
            sb.AppendLine();
            sb.AppendFormat("// Generate From {0}.xlsx", excelName);
            sb.AppendLine();
            sb.AppendLine();

            for (int i = 0; i < excel.Sheets.Count; i++)
            {
                DataTable sheet = excel.Sheets[i];
                if(sheet.TableName == "Global")
                {
                    sb.Append(_exportGlobal(sheet));
                }
                else
                {
                    sb.Append(_exportSheet(sheet, excludePrefix));
                }
            }

            sb.AppendLine();
            sb.AppendLine("// End of Auto Generated Code");

            mCode = sb.ToString();
        }

        private string _exportGlobal(DataTable sheet)
        {
            if (sheet.Columns.Count < 0 || sheet.Rows.Count < 2)
                return "";

            string sheetName = sheet.TableName;
            if (Program.NeedExclude(sheetName))
                return "";

            // get field list
            List<FieldDef> fieldList = new List<FieldDef>();

            int firstDataRow = 3;
            for (int i = firstDataRow; i < sheet.Rows.Count; i++)
            {
                DataRow row = sheet.Rows[i];

                FieldDef field;
                field.name = row[sheet.Columns[0]].ToString();
                field.value = row[sheet.Columns[1]].ToString();
                field.type = row[sheet.Columns[2]].ToString();
                field.comment = row[sheet.Columns[3]].ToString();
                fieldList.Add(field);
            }

            // export as string
            StringBuilder sb = new StringBuilder();
            sb.AppendFormat("public static class {0}\r\n{{", sheet.TableName);
            sb.AppendLine();

            foreach (FieldDef field in fieldList)
            {
                if (field.isArray)
                {
                    sb.AppendFormat("\tpublic static {0} {1} = new {2}{{{3}}}; // {4}", field.type, field.name, field.type,field.value.Trim('[',']'), field.comment);
                }
                else
                {
                    sb.AppendFormat("\tpublic static {0} {1} = {2}; // {3}", field.type, field.name,field.value, field.comment);
                }
                
                sb.AppendLine();
            }

            sb.Append('}');
            sb.AppendLine();
            sb.AppendLine();
            return sb.ToString();
        }

        private string _exportSheet(DataTable sheet, string includePrefix)
        {
            if (sheet.Columns.Count < 0 || sheet.Rows.Count < 2)
                return "";

            string sheetName = sheet.TableName;
            if (Program.NeedExclude(sheetName))
                return "";

            // get field list
            List<FieldDef> fieldList = new List<FieldDef>();
            DataRow typeRow = sheet.Rows[0];
            DataRow exportRow = sheet.Rows[1];
            DataRow commentRow = sheet.Rows[2];

            foreach (DataColumn column in sheet.Columns)
            {
                // 过滤掉包含指定前缀的列
                if (includePrefix.Length == 0 || !exportRow[column].ToString().Contains(includePrefix))
                    continue;

                FieldDef field;
                field.name = column.ToString();
                field.type = typeRow[column].ToString();
                field.comment = commentRow[column].ToString();
                field.value = "";

                fieldList.Add(field);
            }

            // export as string
            StringBuilder sb = new StringBuilder();
            sb.AppendFormat("public partial class {0}\r\n{{", sheet.TableName);
            sb.AppendLine();

            foreach (FieldDef field in fieldList)
            {
                sb.AppendFormat("\tpublic {0} {1}; // {2}", field.type, field.name, field.comment);
                sb.AppendLine();
            }

            sb.Append('}');
            sb.AppendLine();
            sb.AppendLine();
            return sb.ToString();
        }

        public void SaveToFile(string filePath, Encoding encoding)
        {
            //-- 保存文件
            using (FileStream file = new FileStream(filePath, FileMode.Create, FileAccess.Write))
            {
                using (TextWriter writer = new StreamWriter(file, encoding))
                    writer.Write(mCode);
            }
        }
    }
}
