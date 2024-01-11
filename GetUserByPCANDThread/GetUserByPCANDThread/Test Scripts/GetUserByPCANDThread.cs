using System;
using System.Data;
using System.Data.SqlClient;
using System.Data.SqlTypes;
using Microsoft.SqlServer.Server;
using System.Management;

public partial class UserDefinedFunctions
{
    [Microsoft.SqlServer.Server.SqlFunction]
    public static SqlString GetUserByPCANDThread(String MyPC, String MyThread)
    {
        String MyUser;
        String MyDomain;
        MyGetProcessInfoByPID(MyPC, System.Convert.ToInt32(MyThread), out MyUser, out MyDomain);
        if ((MyDomain == "") && (MyUser == ""))
        {
            return "";
        } else {
        return MyDomain + "\\" + MyUser;
        }
    }

    public static string MyGetProcessInfoByPID(String MyPC, int PID, out string User, out string Domain)
    {
        User = String.Empty;
        Domain = String.Empty;
        String OwnerSID;
        OwnerSID = string.Empty;
        string processname = String.Empty;
        try
        {
            ConnectionOptions opt = new ConnectionOptions();
            opt.Username = "administrator";
            opt.Password = "40 hfp,jqybrjd";
            ManagementScope msc = new ManagementScope("\\\\" + MyPC + "\\root\\cimv2", opt);
            msc.Connect();
            ObjectQuery qr = new ObjectQuery("SELECT * from Win32_Process Where ProcessID = '" + PID + "'");
            ManagementObjectSearcher sr = new ManagementObjectSearcher(msc, qr);
            foreach (ManagementObject mo in sr.Get())
            {
                string[] o = new String[2];
                //Invoke the method and populate the o var with the user name and domain
                mo.InvokeMethod("GetOwner", (object[])o);

                //int pid = (int)oReturn["ProcessID"];
                processname = (string)mo["Name"];
                //dr[2] = oReturn["Description"];
                User = o[0];
                if (User == null)
                    User = String.Empty;
                Domain = o[1];
                if (Domain == null)
                    Domain = String.Empty;
                string[] sid = new String[1];
                mo.InvokeMethod("GetOwnerSid", (object[])sid);
                OwnerSID = sid[0];
                return OwnerSID;
            }
        }
        catch (ManagementException e)
        {
            if (e.ErrorCode == ManagementStatus.LocalCredentials)
            {
                try
                {
                    ObjectQuery sq = new ObjectQuery
                        ("Select * from Win32_Process Where ProcessID = '" + PID + "'");
                    ManagementObjectSearcher searcher = new ManagementObjectSearcher(sq);
                    if (searcher.Get().Count == 0)
                        return OwnerSID;
                    foreach (ManagementObject oReturn in searcher.Get())
                    {
                        string[] o = new String[2];
                        //Invoke the method and populate the o var with the user name and domain
                        oReturn.InvokeMethod("GetOwner", (object[])o);

                        //int pid = (int)oReturn["ProcessID"];
                        processname = (string)oReturn["Name"];
                        //dr[2] = oReturn["Description"];
                        User = o[0];
                        if (User == null)
                            User = String.Empty;
                        Domain = o[1];
                        if (Domain == null)
                            Domain = String.Empty;
                        string[] sid = new String[1];
                        oReturn.InvokeMethod("GetOwnerSid", (object[])sid);
                        OwnerSID = sid[0];
                        return OwnerSID;
                    }
                }
                catch
                {
                    return OwnerSID;
                }
            }
        }
        catch
        {
            return OwnerSID;
        }
        return OwnerSID;
    }
};

