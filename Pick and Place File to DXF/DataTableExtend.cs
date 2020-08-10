using System;
using System.Collections.Generic;
using System.Data;
using System.Reflection;
namespace DataTableToDataListExtend
{
    /// <summary>
    /// DataTable扩展方法类
    /// https://blog.csdn.net/pan_junbiao/article/details/82935992
    /// </summary>
    public static class DataTableExtend
    {
        /// <summary>
        /// DataTable转成List
        /// </summary>
        public static List<T> ToDataList<T>(this DataTable dt)
        {
            var list = new List<T>();
            var plist = new List<PropertyInfo>(typeof(T).GetProperties());

            if (dt == null || dt.Rows.Count == 0)
            {
                return null;
            }

            foreach (DataRow item in dt.Rows)
            {
                T s = Activator.CreateInstance<T>();
                for (int i = 0; i < dt.Columns.Count; i++)
                {
                    PropertyInfo info = plist.Find(p => p.Name == dt.Columns[i].ColumnName);
                    if (info != null)
                    {
                        try
                        {
                            if (!Convert.IsDBNull(item[i]))
                            {
                                object v = null;
                                if (info.PropertyType.ToString().Contains("System.Nullable"))
                                {
                                    v = Convert.ChangeType(item[i], Nullable.GetUnderlyingType(info.PropertyType));
                                }
                                else
                                {
                                    v = Convert.ChangeType(item[i], info.PropertyType);
                                }
                                info.SetValue(s, v, null);
                            }
                        }
                        catch (Exception ex)
                        {
                            throw new Exception("字段[" + info.Name + "]转换出错," + ex.Message);
                        }
                    }
                }
                list.Add(s);
            }
            return list;
        }

        /// <summary>
        /// DataTable转成实体对象，仅转行第0行。
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="dt"></param>
        /// <returns></returns>
        public static T ToDataEntity<T>(this DataTable dt)
        {
            T s = Activator.CreateInstance<T>();
            if (dt == null || dt.Rows.Count == 0)
            {
                return default(T);
            }
            var plist = new List<PropertyInfo>(typeof(T).GetProperties());
            for (int i = 0; i < dt.Columns.Count; i++)
            {
                PropertyInfo info = plist.Find(p => p.Name == dt.Columns[i].ColumnName);
                if (info != null)
                {
                    try
                    {
                        if (!Convert.IsDBNull(dt.Rows[0][i]))
                        {
                            object v = null;
                            if (info.PropertyType.ToString().Contains("System.Nullable"))
                            {
                                v = Convert.ChangeType(dt.Rows[0][i], Nullable.GetUnderlyingType(info.PropertyType));
                            }
                            else
                            {
                                v = Convert.ChangeType(dt.Rows[0][i], info.PropertyType);
                            }
                            info.SetValue(s, v, null);
                        }
                    }
                    catch (Exception ex)
                    {
                        throw new Exception("字段[" + info.Name + "]转换出错," + ex.Message);
                    }
                }
            }
            return s;
        }

        /// <summary>
        /// DataTable转成实体对象，指定行转换。
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="dt"></param>
        /// <param name="RowIndex">行索引</param>
        /// <returns></returns>
        public static T ToDataEntity<T>(this DataTable dt, int RowIndex)
        {
            T s = Activator.CreateInstance<T>();
            if (dt == null || dt.Rows.Count == 0)
            {
                return default(T);
            }
            var plist = new List<PropertyInfo>(typeof(T).GetProperties());
            for (int i = 0; i < dt.Columns.Count; i++)
            {
                PropertyInfo info = plist.Find(p => p.Name == dt.Columns[i].ColumnName);
                if (info != null)
                {
                    if (RowIndex > dt.Rows.Count - 1 || RowIndex < 0)
                    {
                        throw new Exception("RowIndex Error");
                    }
                    try
                    {
                        if (!Convert.IsDBNull(dt.Rows[0][i]))
                        {
                            object v = null;
                            if (info.PropertyType.ToString().Contains("System.Nullable"))
                            {
                                v = Convert.ChangeType(dt.Rows[RowIndex][i], Nullable.GetUnderlyingType(info.PropertyType));
                            }
                            else
                            {
                                v = Convert.ChangeType(dt.Rows[RowIndex][i], info.PropertyType);
                            }
                            info.SetValue(s, v, null);
                        }
                    }
                    catch (Exception ex)
                    {
                        throw new Exception("字段[" + info.Name + "]转换出错," + ex.Message);
                    }
                }
            }
            return s;
        }

        /// <summary>
        /// List转成DataTable
        /// </summary>
        /// <typeparam name="T">实体类型</typeparam>
        /// <param name="entities">实体集合</param>
        public static DataTable ToDataTable<T>(List<T> entities)
        {
            if (entities == null || entities.Count == 0)
            {
                return null;
            }

            var result = CreateTable<T>();
            FillData(result, entities);
            return result;
        }

        /// <summary>
        /// 创建表
        /// </summary>
        private static DataTable CreateTable<T>()
        {
            var result = new DataTable();
            var type = typeof(T);
            foreach (var property in type.GetProperties(System.Reflection.BindingFlags.Public | System.Reflection.BindingFlags.Instance))
            {
                var propertyType = property.PropertyType;
                if ((propertyType.IsGenericType) && (propertyType.GetGenericTypeDefinition() == typeof(Nullable<>)))
                    propertyType = propertyType.GetGenericArguments()[0];
                result.Columns.Add(property.Name, propertyType);
            }
            return result;
        }

        /// <summary>
        /// 填充数据
        /// </summary>
        private static void FillData<T>(DataTable dt, IEnumerable<T> entities)
        {
            foreach (var entity in entities)
            {
                dt.Rows.Add(CreateRow(dt, entity));
            }
        }

        /// <summary>
        /// 创建行
        /// </summary>
        private static DataRow CreateRow<T>(DataTable dt, T entity)
        {
            DataRow row = dt.NewRow();
            var type = typeof(T);
            foreach (var property in type.GetProperties(System.Reflection.BindingFlags.Public | System.Reflection.BindingFlags.Instance))
            {
                row[property.Name] = property.GetValue(entity) ?? DBNull.Value;
            }
            return row;
        }
    }
}


//Example
/*
    private void button3_Click(object sender, EventArgs e)
    {
        //创建一个DataTable对象
        DataTable dt = CreateDataTable();
        //1、DataTable转成List
        List<UserInfo> userList = dt.ToDataList<UserInfo>();
        //2、DataTable转成实体对象， 注意仅 第0 行
        UserInfo user = dt.ToDataEntity<UserInfo>();
        //3、DataTable转成实体对象， 指定行
        UserInfo user2 = dt.ToDataEntity<UserInfo>(1);
        //4、List转成DataTable
        DataTable dt2 = DataTableExtend.DataTableExtend.ToDataTable<UserInfo>(userList);
    }
    /// <summary>
    /// 创建DataTable对象
    /// </summary>
    public static DataTable CreateDataTable()
    {
        //创建DataTable
        DataTable dt = new DataTable("NewDt");
        //创建自增长的ID列
        DataColumn dc = dt.Columns.Add("ID", Type.GetType("System.Int32"));
        dc.AutoIncrement = true;   //自动增加
        dc.AutoIncrementSeed = 1;  //起始为1
        dc.AutoIncrementStep = 1;  //步长为1
        dc.AllowDBNull = false;    //非空
        //创建其它列表
        dt.Columns.Add(new DataColumn("Name", Type.GetType("System.String")));
        dt.Columns.Add(new DataColumn("Age", Type.GetType("System.Int32")));
        dt.Columns.Add(new DataColumn("Score", Type.GetType("System.Decimal")));
        dt.Columns.Add(new DataColumn("CreateTime", Type.GetType("System.DateTime")));
        //创建数据
        DataRow dr = dt.NewRow();
        dr["Name"] = "张三";
        dr["Age"] = 28;
        dr["Score"] = 85.5;
        dr["CreateTime"] = DateTime.Now;
        dt.Rows.Add(dr);
        dr = dt.NewRow();
        dr["Name"] = "李四";
        dr["Age"] = 24;
        dr["Score"] = 72;
        dr["CreateTime"] = DateTime.Now;
        dt.Rows.Add(dr);
        dr = dt.NewRow();
        dr["Name"] = "王五";
        dr["Age"] = 36;
        dr["Score"] = 63.5;
        dr["CreateTime"] = DateTime.Now;
        dt.Rows.Add(dr);
        return dt;
    }
    /// <summary>
    /// 用户信息类
    /// </summary>
    public class UserInfo
    {
        /// <summary>
        /// 编号
        /// </summary>
        public int ID { get; set; }
        /// <summary>
        /// 名称
        /// </summary>
        public string Name { get; set; }
        /// <summary>
        /// 年龄
        /// </summary>
        public int Age { get; set; }
        /// <summary>
        /// 成绩
        /// </summary>
        public double Score { get; set; }
        /// <summary>
        /// 创建时间
        /// </summary>
        public DateTime? CreateTime { get; set; }
    }
 
 */
