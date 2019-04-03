using System;
using System.Data;
using System.IO;
using ExcelDataReader;
using System.Diagnostics;
using System.Collections.Generic;

namespace ExcelToLua
{
    class Program
    {
        static string inputPath = "";
        static string outPath = "";
        static string ThemeName = "ThemeJsonData";
        static List<string> mListColumnTypeInfo = null;
        static List<int> mListValidColumn = null;

        static void Main(string[] args)
        {
            System.Text.Encoding.RegisterProvider(System.Text.CodePagesEncodingProvider.Instance);

            string filePath = inputPath + ThemeName + ".xls";
            using (var stream = File.Open(filePath, FileMode.Open, FileAccess.Read))
            {
                using (var reader = ExcelReaderFactory.CreateReader(stream))
                {
                    do
                    {
                        while (reader.Read())
                        {

                        }
                    } while (reader.NextResult());

                    System.Data.DataSet result = reader.AsDataSet();
                    ParseDataSet(result);
                }
            }
        }

        static void ParseDataSet(DataSet result)
        {
            ParseColumnType(result);
            ParseData(result);
        }

        static void ParseColumnType(DataSet result)
        {
            mListColumnTypeInfo = new List<string>();
            mListValidColumn = new List<int>();
            for (int i = 0; i < result.Tables.Count; i++)
            {
                DataTable mTable = result.Tables[i];
                if (mTable.TableName == "Export Summary")
                {
                    continue;
                }

                for (int j = 0; j < mTable.Columns.Count; j++)
                {
                    string fieldName = string.Empty;
                    string desName = string.Empty;
                    string typeName = string.Empty;
                    string subtypeName = string.Empty;

                    for (int k = 0; k < mTable.Rows.Count; k++)
                    {
                        string value = mTable.Rows[k][j].ToString();
                        //Console.WriteLine("(" + j +", " + k + ") : " + value);

                        if (k == 1) //字段名
                        {
                            fieldName = value;
                            Console.WriteLine("字段名： " + fieldName + " | " + value);
                        }
                        else if (k == 2) //类型
                        {
                            typeName = value;
                            if (typeName.StartsWith("TABLE", StringComparison.Ordinal))
                            {
                                string[] typeNameList = typeName.Split(",");
                                typeName = typeNameList[0];
                                subtypeName = typeNameList[1];
                            }

                            Console.WriteLine("类型名： " + typeName + " | " + value);
                        }
                        else if (k == 3) //描述
                        {
                            desName = value;
                        }
                        else
                        {
                            continue;
                        }
                    }

                    mListColumnTypeInfo.Add(fieldName);
                    mListColumnTypeInfo.Add(desName);
                    mListColumnTypeInfo.Add(typeName);
                    mListColumnTypeInfo.Add(subtypeName);
                }
            }
        }

        static void CheckValidColumn()
        {

        }

        static void ParseData(DataSet result)
        {
            string outStr = "local " + ThemeName + " = {\n";

            for (int i = 0; i < result.Tables.Count; i++)
            {
                DataTable mTable = result.Tables[i];
                if (mTable.TableName == "Export Summary")
                {
                    continue;
                }

                Console.WriteLine("TableName: " + mTable.TableName);

                outStr += "\t" + mTable.TableName + " = {\n";

                for (int j = 4; j < mTable.Rows.Count; j++)
                {
                    outStr += "\t\t{";
                    for (int k = 0; k < mTable.Columns.Count; k++)
                    {
                        string value = mTable.Rows[j][k].ToString();

                        string fieldName = mListColumnTypeInfo[k * 4 + 0];
                        string desName = mListColumnTypeInfo[k * 4 + 1];
                        string typeName = mListColumnTypeInfo[k * 4 + 2];
                        string subtypeName = mListColumnTypeInfo[k * 4 + 3];

                        if (fieldName == string.Empty || typeName == string.Empty)
                        {
                            continue;
                        }

                        if (typeName == "FLOAT")
                        {
                            float fValue = float.Parse(value);
                            outStr += fieldName + " = " + fValue;
                        }
                        else if (typeName == "INT")
                        {
                            int nValue = int.Parse(value);
                            outStr += fieldName + " = " + nValue;
                        }
                        else if (typeName == "STRING")
                        {
                            string strValue = value;
                            outStr += fieldName + " = \"" + strValue + "\"";
                        }
                        else if (typeName == "TABLE")
                        {
                            string strValue = value;
                            if (strValue.StartsWith("{", StringComparison.Ordinal))
                            {
                                strValue = strValue.Substring(1, strValue.Length - 1);
                            }

                            if (strValue.EndsWith("}", StringComparison.Ordinal))
                            {
                                strValue = strValue.Substring(0, strValue.Length - 1);
                            }

                            string[] words = strValue.Split(',');

                            outStr += fieldName + " = {";
                            for (int n = 0; n < words.Length; n++)
                            {
                                string tempvalue = words[n];
                                if (subtypeName == "FLOAT")
                                {
                                    float fValue = float.Parse(tempvalue);
                                    if (n == 0)
                                    {
                                        outStr += fValue;
                                    }
                                    else
                                    {
                                        outStr += ", " + fValue;
                                    }
                                }
                                else if (subtypeName == "INT")
                                {
                                    int nValue = int.Parse(tempvalue);
                                    if (n == 0)
                                    {
                                        outStr += nValue;
                                    }
                                    else
                                    {
                                        outStr += ", " + nValue;
                                    }
                                }
                                else if (subtypeName == "STRING")
                                {
                                    if (n == 0)
                                    {
                                        outStr += tempvalue;
                                    }
                                    else
                                    {
                                        outStr += ", " + tempvalue;
                                    }
                                }
                                else
                                {
                                    Debug.Assert(false);
                                }
                            }
                            outStr += "}";
                        }
                        else
                        {
                            Debug.Assert(false);
                        }

                        if (k < mTable.Columns.Count - 1)
                        {
                            outStr += ",\t";
                        }
                    }

                    outStr += "},\n";
                }

                outStr += "\t},\n\n";
            }

            outStr += "}\n\n";

            outStr += "return " + ThemeName + "\n";

            Console.WriteLine("outStr" + outStr);

            String outFileName = outPath + ThemeName + ".lua";
            File.WriteAllText(outFileName, outStr);
        }
    }
}
