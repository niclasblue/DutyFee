using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DutyFee
{   
    /// <summary>
    /// 从json配置文件中读取节假日信息，从而判断当前所给日期参数的类型
    /// </summary>
    public class HolidayChecker
    {
        public HolidayChecker(string configPath)
        {
            ConfigPath = configPath;
            cacheDateList = GetConfigList(ConfigPath);
        }

        private List<DateModel> cacheDateList { get; set; }
        public string ConfigPath { get; set; }

        #region 私有方法
        /// <summary>
        /// 读取文件
        /// </summary>
        /// <param name="filePath"></param>
        /// <returns>文件内容</returns>
        private string GetFileContent(string filePath)
        {
            string result = "";
            if (File.Exists(filePath))
            {
                result = File.ReadAllText(filePath);
            }
            return result;
        }
        /// <summary>
        /// 获取配置的Json文件
        /// </summary>
        /// <returns>经过反序列化之后的对象集合</returns>
        private List<DateModel> GetConfigList(string path)
        {
            string fileContent = GetFileContent(path);
            if (!string.IsNullOrWhiteSpace(fileContent))
            {
                cacheDateList = Newtonsoft.Json.JsonConvert.DeserializeObject<List<DateModel>>(fileContent);
            }
            return cacheDateList;
        }
        /// <summary>
        /// 获取指定年份的数据
        /// </summary>
        /// <param name="year"></param>
        /// <returns></returns>
        private DateModel GetConfigDataByYear(int year)
        {
            if (cacheDateList == null)//取配置数据
                GetConfigList(ConfigPath);
            DateModel result = cacheDateList.FirstOrDefault(m => m.Year == year);
            return result;
        }

        #endregion

        #region 公共方法
        /// <summary>
        /// 判断日期类型
        /// </summary>
        /// <param name="currDate">要判断的日期</param>
        /// <param name="thisYearData">当前的数据</param>
        /// <returns>日期类型（工作日、周末还是节假日）</returns>
        public DateType CheckDayType(DateTime currDate)
        {

            DateModel thisYearData = GetConfigDataByYear(currDate.Year);


            string date = currDate.ToString("MMdd");
            int week = (int)currDate.DayOfWeek;

            if (thisYearData.Work.Contains(date))
            {
                return DateType.Workday;
            }

            if (thisYearData.Holiday.Contains(date))
            {
                return DateType.Holiday;
            }

            if (week != 0 && week != 6)
            {
                return DateType.Workday;
            }
            return DateType.Weekend;
        }

        public void CleraCacheData()
        {
            if (cacheDateList != null)
            {
                cacheDateList.Clear();
            }
        }
        #endregion
    }
}
