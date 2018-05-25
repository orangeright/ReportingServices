using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace ReportingServices
{
    public static class Parameters
    {
        public static Pleasanter Pleasanter;
    }

    public class Pleasanter
    {
        public string ApiKey;
        public string Uri;
        public string TemplatePath;
    }

    public static class JsonData
    {
        //public static Dictionary<string,string> Jdata;

        public static string Jdata = @"{
  ""UpdatedTime"": ""2018-05-18T11:34:18"",
  ""ResultId"": 269,
  ""Ver"": 2,
  ""Owner"": 12,
  ""Class001"": ""田口一郎"",
  ""Class002"": ""タグチイチロウ"",
  ""Class003"": ""男"",
  ""Class004"": ""○○大学"",
  ""Class005"": ""埼玉高速鉄道線"",
  ""Class006"": ""浦和美園駅"",
  ""Class007"": ""第二種情報処理技術者"",
  ""Class008"": ""1997年12月"",
  ""Class009"": ""ソフトウェア開発技術者"",
  ""Class010"": ""2003年6月"",
  ""Class011"": """",
  ""Class012"": """",
  ""Class013"": """",
  ""Class014"": """",
  ""Class015"": """",
  ""Class016"": """",
  ""Class017"": ""◎"",
  ""Class018"": ""◎"",
  ""Class019"": ""◎"",
  ""Class020"": ""◎"",
  ""Class021"": ""◎"",
  ""Class022"": ""△"",
  ""Class023"": ""△"",
  ""Class024"": ""◎"",
  ""Class025"": ""◎"",
  ""Class026"": ""◎"",
  ""Class027"": ""△"",
  ""NumA"": 45,
  ""NumB"": 1973,
  ""NumC"": 2,
  ""NumD"": 17,
  ""Num001"": 4,
  ""Num002"": 3,
  ""Num003"": 3,
  ""Num004"": 0,
  ""Num005"": 0,
  ""Num006"": 0,
  ""Num007"": 4,
  ""Num008"": 4,
  ""Num009"": 4,
  ""Num010"": 3,
  ""Num011"": 3,
  ""Num012"": 0,
  ""Num013"": 2,
  ""Num014"": 4,
  ""Num015"": 2,
  ""Num016"": 3,
  ""Num017"": 1,
  ""Num018"": 0,
  ""Num019"": 1,
  ""Num020"": 0,
  ""Num021"": 3,
  ""Num022"": 4,
  ""Num023"": 1,
  ""Num024"": 0,
  ""Num025"": 0,
  ""Num026"": 0,
  ""Num027"": 0,
  ""Num028"": 1,
  ""Num029"": 0,
  ""Num030"": 0,
  ""Num031"": 0,
  ""Num032"": 4,
  ""Num033"": 4,
  ""Num034"": 3,
  ""Num035"": 0,
  ""Num036"": 0,
  ""Num037"": 0,
  ""Num038"": 0,
  ""Num039"": 3,
  ""Num040"": 0,
  ""Num041"": 3,
  ""Num042"": 1,
  ""Num043"": 3,
  ""Num044"": 0,
  ""Num045"": 0,
  ""Num046"": 0,
  ""Comments"": ""[]"",
  ""Updator"": 12,
  ""PrintDate"": ""2018/5/18"",
  ""PrintTime"": ""16:50:11""
}
";
    }

}