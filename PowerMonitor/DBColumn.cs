
using System;
using System.Collections.Generic;
using System.Text;

namespace PowerMonitor
{
    /// <summary>
    /// 数据库字段对应属性类
    /// 说明    ：数据库字段对应属性类<br/>
    /// 作者    ：niu<br/>
    /// 创建时间：2011-07-21<br/>
    /// 最后修改：2011-07-21<br/>
    /// </summary>
    [AttributeUsage(AttributeTargets.Property|AttributeTargets.Method,AllowMultiple=true,Inherited=false)]
    public class DBColumn : Attribute
    {
        private string _colName;
        /// <summary>
        /// 数据库字段
        /// </summary>
        public string ColName
        { 
            get { return _colName; }
            set { _colName = value; }
        }
        /*
 AttributeTargets 枚举 
 成员名称 说明 
 All 可以对任何应用程序元素应用属性。  
 Assembly 可以对程序集应用属性。  
 Class 可以对类应用属性。  
 Constructor 可以对构造函数应用属性。  
 Delegate 可以对委托应用属性。  
 Enum 可以对枚举应用属性。  
 Event 可以对事件应用属性。  
 Field 可以对字段应用属性。  
 GenericParameter 可以对泛型参数应用属性。  
 Interface 可以对接口应用属性。  
 Method 可以对方法应用属性。  
 Module 可以对模块应用属性。 注意 
 Module 指的是可移植的可执行文件（.dll 或 .exe），而非 Visual Basic 标准模块。
 Parameter 可以对参数应用属性。  
 Property 可以对属性 (Property) 应用属性 (Attribute)。  
 ReturnValue 可以对返回值应用属性。  
 Struct 可以对结构应用属性，即值类型 
     */

        /*
         这里会有四种可能的组合： 
   
  [AttributeUsage(AttributeTargets.Class, AllowMultiple = false, Inherited = false ] 
  [AttributeUsage(AttributeTargets.Class, AllowMultiple = true, Inherited = false ] 
  [AttributeUsage(AttributeTargets.Class, AllowMultiple = false, Inherited = true ] 
  [AttributeUsage(AttributeTargets.Class, AllowMultiple = true, Inherited = true ]

　　第一种情况： 

　　如果我们查询Derive类，我们将会发现Help特性并不存在，因为inherited属性被设置为false。

　　第二种情况：

　　和第一种情况相同，因为inherited也被设置为false。

　　第三种情况：

　　为了解释第三种和第四种情况，我们先来给派生类添加点代码： 
   
  [Help("BaseClass")] 
  public class Base 
  { 
  } 
  [Help("DeriveClass")] 
  public class Derive : Base 
  { 
  }

　　现在我们来查询一下Help特性，我们只能得到派生类的属性，因为inherited被设置为true，但是AllowMultiple却被设置为false。因此基类的Help特性被派生类Help特性覆盖了。

　　第四种情况：

　　在这里，我们将会发现派生类既有基类的Help特性，也有自己的Help特性，因为AllowMultiple被设置为true。
         */
    }
    
}