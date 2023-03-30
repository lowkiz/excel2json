using System;
using System.IO;
using System.Data;
using System.Text;
using System.Collections.Generic;


namespace excel2json
{
    class GoDefineGenerator
    {
        struct FieldDef
        {
            public string name;
            public string comment;
        }

        string mCode;

        public string code
        {
            get
            {
                return this.mCode;
            }
        }

        public GoDefineGenerator(string excelName, ExcelLoader excel, string excludePrefix)
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
            sb.AppendLine("package generated");
            sb.AppendLine();

            for (int i = 0; i < excel.Sheets.Count; i++)
            {
                DataTable sheet = excel.Sheets[i];
                sb.Append(_exportSheet(sheet, excludePrefix));
            }

            sb.AppendLine();
            sb.AppendLine("// End of Auto Generated Code");

            mCode = sb.ToString();
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
            DataRow exportRow = sheet.Rows[1];
            DataRow commentRow = sheet.Rows[2];

            foreach (DataColumn column in sheet.Columns)
            {
                if (includePrefix.Length == 0 || !exportRow[column].ToString().Contains(includePrefix))
                    continue;

                FieldDef field;
                field.name = column.ToString();
                field.comment = commentRow[column].ToString();

                fieldList.Add(field);
            }

            // export as string
            StringBuilder sb = new StringBuilder();
            sb.AppendFormat("type {0} struct {{", sheet.TableName+"Data");
            sb.AppendLine();

            foreach (FieldDef field in fieldList)
            {
                sb.AppendFormat("\t{0} {1} // {2}", field.name, "string", field.comment);
                sb.AppendLine();
            }

            sb.Append('}');
            sb.AppendLine();
            sb.AppendLine();
            sb.Append(_exportSheetVar(sheet.TableName, fieldList));
            return sb.ToString();
        }

        private string _exportSheetVar(string sheetName, List<FieldDef> fieldList)
        {
            StringBuilder sb = new StringBuilder();
            sb.AppendLine("var (");
            sb.AppendFormat("\t{0} = &{1}{{", sheetName, sheetName+"Data");
            sb.AppendLine();
            foreach (FieldDef field in fieldList)
            {
                sb.AppendFormat("\t\t{0}: \"{1}\",", field.name, field.name);
                sb.AppendLine();
            }
            sb.Append('}');
            sb.AppendLine();
            sb.AppendLine(")");
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
