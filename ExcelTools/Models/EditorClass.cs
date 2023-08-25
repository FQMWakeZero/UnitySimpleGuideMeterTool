using System.Collections;
using System.Collections.Generic;
using UnityEngine;


namespace Framework
{
    public class EditorClass
    {
        private static EditorClass _Instance;
        public static EditorClass Instance { get { _Instance ??= new EditorClass(); return _Instance; } }
        public EditorClass()
        {
            ClassName = null;
            Field = new List<EditorClass_Field>();
        }

        public string ClassName { get; set; }               //类名
        public List<EditorClass_Field> Field { get; set; }  //字段表
    }

    public class EditorClass_Field
    {
        public string FieldName { get; set; }               //字段名
        public string FieldComment { get; set; }            //字段注释
        public string FieldType { get; set; }               //字段类型
    }
}

