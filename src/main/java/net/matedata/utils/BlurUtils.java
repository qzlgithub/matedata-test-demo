package net.matedata.utils;

public class BlurUtils
{
    private BlurUtils()
    {
    }

    /**
     * 用户身份证号模糊化
     */
    public static String blurPersonIdNo(String idNo)
    {
        if(idNo == null || idNo.length() < 8)
        {
            return idNo == null ? "" : idNo;
        }
        return idNo.substring(0, 4) + "**********" + idNo.substring(idNo.length() - 4);
    }

    /**
     * 用户手机号模糊化
     */
    public static String blurPersonMobile(String mobile)
    {
        if(mobile == null || mobile.length() < 7)
        {
            return mobile == null ? "" : mobile;
        }
        return mobile.substring(0, 3) + "****" + mobile.substring(mobile.length() - 4);
    }
}
