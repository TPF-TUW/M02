using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

public class Config
{
    public const string PATH_FILE_CONNECTION = @"\\172.16.0.190\MDS_Project\FileConfig\Configue.ini";
    public const string CONNECTION_SECTION = "ConnectionString";
    public static string DEF_SERVER { get; set; }
    public static string DEF_DB { get; set; }
    public static string DEF_USER { get; set; }
    public static string DEF_PW { get; set; }


}
