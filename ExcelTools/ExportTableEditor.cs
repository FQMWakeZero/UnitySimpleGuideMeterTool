/// Unity导表工具，用法，读取Assets/Excels下的表格，目前支持xlsx，xls，csv，xlsm四种类型，数据格式支持int，int[]，string,string[],fload,fload[]六种类型
/// 表格排版格式为：第一行不读取，第二行为变量名，第三行是类型名，第四行是备注也是注释，第五行以下包括第五行均为数据集
/// C#实体类默认保存在Assets/Excels/Models目录下，json数据默认报错在Assets/Resources/Configs目录下
/// 如果遇到问题可以联系QQ：2251313670
/// 代码可任意使用，但务必保留顶部注释，代码禁止售卖
/// 代码所使用的额外库：ExcelDataReader，Newtonsoft.json
/// 开发环境：Unity2021.3.27f1c2 Vs 2022
/// 理论来说变量名可以用中文，代码最底部的字典可以根据规则自行添加，额外的类型解析可以自行补充
/// 如果代码报错，请自行为项目添加Newtonsoft.json库

using UnityEditor;
using UnityEngine;
using ExcelDataReader;
using System.IO;
using System;
using Newtonsoft.Json;
using System.Reflection;
using System.Reflection.Emit;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace Framework
{
    public class ExportTableEditor
    {
        public static char[] Separator = new char[] { ';', ',', '，', '；' };
        public static string _namespace = "GameModel";

        [MenuItem("Tools/导表菜单/导出")]
        public static void PlayExport()
        {
            var path = $"{Application.dataPath}/Excels";
            if (!System.IO.Directory.Exists(path))
            {
                Debug.Log($"指定目录不存在，请将Excel文件放在Assets/Excels文件夹目录下");
                try { System.IO.Directory.CreateDirectory(path); }
                catch (System.Exception e)
                {
                    Debug.Log($"创建目录：{path} 失败，错误：{e}");
                    throw;
                }
                return;
            }

            FileStream fileStream = null;
            foreach (string extension in new string[] { "*.xlsx", "*.xls", "*.csv", "*.xlsm" })
                foreach (string filePath in System.IO.Directory.GetFiles(path, extension))
                {

                    IExcelDataReader excelDataReader = null;
                    try
                    {
                        fileStream = new FileStream(filePath, FileMode.Open, FileAccess.Read);

                        excelDataReader = ExcelReaderFactory.CreateReader(fileStream);
                    }
                    catch (Exception) { }

                    if (excelDataReader == null)
                    {
                        Debug.Log($"{filePath} 文件没有正常打开，是否未关闭表？");
                        fileStream.Close();
                        return;
                    }

                    if (excelDataReader.IsClosed)
                    {
                        Debug.Log($"{filePath} 文件已关闭无法读取");
                        fileStream.Close();
                        return;
                    }

                    do //读取多个页面
                    {
                        int Count = 0;
                        EditorClass editorClass = new EditorClass();
                        editorClass.ClassName = excelDataReader.Name;
                        List<object> dynamicObjects = null;
                        Type dynamicType = null;


                        while (excelDataReader.Read())  //从行开始读取
                        {
                            if (excelDataReader[0] == null) continue;

                            for (int column = 0; column < excelDataReader.FieldCount; column++)
                            {
                                if (excelDataReader[column] == null) continue;

                                switch (Count)
                                {
                                    case 1:
                                        editorClass.Field.Add(new EditorClass_Field() { FieldName = excelDataReader[column].ToString() });
                                        break;
                                    case 2:
                                        editorClass.Field[column].FieldType = excelDataReader[column].ToString();
                                        break;
                                    case 3:
                                        editorClass.Field[column].FieldComment = excelDataReader[column].ToString();
                                        break;
                                }
                            }

                            if (Count == 3)
                            {
                                //开始处理实体类
                                dynamicType = Code(editorClass);
                                dynamicObjects = new List<object>();
                            }

                            if (Count >= 4)
                            {
                                object dynamicObject = Activator.CreateInstance(dynamicType);

                                for (int column = 0; column < excelDataReader.FieldCount; column++)
                                {
                                    if (excelDataReader[column] == null)
                                    {
                                        Debug.Log($"有一处表格为空，但理论上不应该为空，请检查{filePath}行：{Count} 列：{column}");
                                        continue;
                                    }

                                    FieldInfo propertyInfo = dynamicType.GetField(editorClass.Field[column].FieldName);
                                    try
                                    {
                                        propertyInfo.SetValue(dynamicObject, ChangeType(excelDataReader[column].ToString(), editorClass.Field[column].FieldType));
                                    }
                                    catch (Exception e) { Debug.Log($"表格读取错误，行：{Count} 列：{column} 错误数据：{excelDataReader[column]} 错误内容：{e}"); fileStream.Close(); throw; }

                                }
                                dynamicObjects.Add(dynamicObject);
                            }
                            Count++;
                        }

                        System.IO.Directory.CreateDirectory($"{Application.dataPath}/Excels/Models");
                        System.IO.Directory.CreateDirectory($"{Application.dataPath}/Resources/Configs");

                        var Csharp = GenerateCSharpCode(editorClass, Path.Combine(Application.dataPath, "Resources", "Config"));
                        var Json = JsonConvert.SerializeObject(dynamicObjects, Formatting.Indented);

                        try
                        {
                            System.IO.File.WriteAllText($"{Application.dataPath}/Excels/Models/{editorClass.ClassName}.cs", Csharp, Encoding.UTF8);
                            System.IO.File.WriteAllText($"{Application.dataPath}/Resources/Configs/{editorClass.ClassName}.json", Json, Encoding.UTF8);
                        }
                        catch (Exception e)
                        {
                            Debug.Log($"{filePath}文件写入失败，Json {Json} ; C# {Csharp} 报错信息：{e}");
                            throw;
                        }


                        Debug.Log($"json对象以及实体类已创建完毕:{filePath}");
                    } while (excelDataReader.NextResult());
                    excelDataReader.Close();
                }
            fileStream.Close();

            AssetDatabase.Refresh();
        }


        private static object ChangeType(string data, string type)
        {

            Type itemType = GetType(type);
            if (data == "null")
                return null;
            else if (itemType == typeof(int))
                return int.Parse(data);
            else if (itemType == typeof(int[]))
                return data.Split(Separator).Select(int.Parse).ToArray();
            else if (itemType == typeof(string[]))
                return data.Split(Separator, StringSplitOptions.RemoveEmptyEntries);
            else if (itemType == typeof(string))
                return data;
            else if (itemType == typeof(float))
                return float.Parse(data);
            else if (itemType == typeof(float[]))
                return data.Split(Separator, StringSplitOptions.RemoveEmptyEntries).Select(s => float.Parse(s.Trim())).ToArray();
            else if (itemType == typeof(bool))
                return bool.Parse(data);
            else
            {
                Debug.Log($"字段：{data} 不是一个有效的类型,当前未添加此类型的解析");
                return default;
            }
        }

        private static string GenerateCSharpCode(EditorClass editorClass, string path)
        {
            StringBuilder codeBuilder = new StringBuilder();
            Type type = Code(editorClass);

            codeBuilder.AppendLine("using Newtonsoft.Json;");
            codeBuilder.AppendLine("using UnityEngine;");
            codeBuilder.AppendLine("using System;");
            codeBuilder.AppendLine("using System.Collections.Generic;");
            codeBuilder.AppendLine($"//这是自动生成的代码，请勿修改");
            codeBuilder.AppendLine($"namespace {_namespace}");
            codeBuilder.AppendLine($"{{");

            codeBuilder.AppendLine($"   public class {type.Name}");
            codeBuilder.AppendLine("    {");

            codeBuilder.AppendLine($"       public {type.Name}()");
            codeBuilder.AppendLine($"       {{");
            codeBuilder.AppendLine($"           root = new List<{type.Name}Subobject>();");
            codeBuilder.AppendLine($"       }}");

            codeBuilder.AppendLine($"       public static {type.Name} Get()");
            codeBuilder.AppendLine($"       {{");
            codeBuilder.AppendLine($"           {type.Name} {type.Name} = new {type.Name}();");
            codeBuilder.AppendLine($"           {type.Name}.root = JsonConvert.DeserializeObject<List<{type.Name}Subobject>>(Resources.Load<TextAsset>(\"Configs\\\\{type.Name}\").text);");
            codeBuilder.AppendLine($"           return {type.Name};");
            codeBuilder.AppendLine($"       }}");

            codeBuilder.AppendLine($"       public {type.Name}Subobject find(string ID)");
            codeBuilder.AppendLine($"       {{");
            codeBuilder.AppendLine($"           foreach (var item in root)");
            codeBuilder.AppendLine($"               if (item.ID == ID)");
            codeBuilder.AppendLine($"                   return item;");
            codeBuilder.AppendLine($"           return null;");
            codeBuilder.AppendLine($"       }}");

            codeBuilder.AppendLine($"       public List<{type.Name}Subobject> root {{ get; set; }}");

            codeBuilder.AppendLine($"   }}");



            codeBuilder.AppendLine($"   public class {type.Name}Subobject");
            codeBuilder.AppendLine($"   {{");

            var Fields = type.GetFields();
            for (int i = 0; i < Fields.Length; i++)
            {
                codeBuilder.AppendLine($"       /// <summary>");
                codeBuilder.AppendLine($"       /// {editorClass.Field[i].FieldComment}");
                codeBuilder.AppendLine($"       /// </summary>");
                codeBuilder.AppendLine($"       public {Fields[i].FieldType.Name} {Fields[i].Name} {{ get; set; }}");
            }

            codeBuilder.AppendLine($"   }}");
            codeBuilder.AppendLine($"}}");

            return codeBuilder.ToString();
        }


        /// <summary>
        /// 动态生成一份实体类
        /// </summary>
        private static Type Code(EditorClass data)
        {

            AssemblyName assemblyName = new AssemblyName("DynamicEntities");
            AssemblyBuilder assemblyBuilder = AssemblyBuilder.DefineDynamicAssembly(assemblyName, AssemblyBuilderAccess.Run);
            ModuleBuilder moduleBuilder = assemblyBuilder.DefineDynamicModule("MainModule");
            TypeBuilder typeBuilder = moduleBuilder.DefineType(data.ClassName, TypeAttributes.Public | TypeAttributes.Class);

            for (int i = 0; i < data.Field.Count; i++)
            {
                typeBuilder.DefineField(data.Field[i].FieldName, GetType(data.Field[i].FieldType), FieldAttributes.Public);
            }

            return typeBuilder.CreateType();
        }

        private static Type GetType(string type)
        {
            foreach (var item in TypeDirectory)
                foreach (var name in item.Key)
                    if (type == name)
                        return Type.GetType(item.Value);
            return typeof(object);

        }

        private static Dictionary<string[], string> TypeDirectory = new Dictionary<string[], string>
        {
            { new string[]{ "string" , "String" } , "System.String" },
            { new string[]{ "int" , "number" } , "System.Int32" },
            { new string[]{ "ints" , "int_array" } , "System.Int32[]" },
            { new string[]{ "float" } , "System.Single" },
            { new string[]{ "floats" } , "System.Single[]" },
            { new string[]{ "strings" , "string_array" } , "System.String[]" },
            { new string[]{ "bool" , "Bool" } , "System.Boolean" },
            { new string[]{ "null" , "Null" } , "System.Nullable" },
        };
    }
}


