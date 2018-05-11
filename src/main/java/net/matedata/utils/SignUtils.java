package net.matedata.utils;

import com.alibaba.fastjson.JSON;

import javax.crypto.Mac;
import javax.crypto.spec.SecretKeySpec;
import java.io.UnsupportedEncodingException;
import java.security.InvalidKeyException;
import java.security.NoSuchAlgorithmException;
import java.util.HashMap;
import java.util.Map;
import java.util.SortedMap;

public class SignUtils
{
    private static final String HMAC_SHA_256 = "HmacSHA256";
    private static final String UTF_8 = "UTF-8";

    private SignUtils()
    {
    }

    public static String sign(String content, String key)
            throws NoSuchAlgorithmException, UnsupportedEncodingException, InvalidKeyException
    {
        Mac hmacSHA256 = Mac.getInstance(HMAC_SHA_256);
        SecretKeySpec spec = new SecretKeySpec(key.getBytes(), HMAC_SHA_256);
        hmacSHA256.init(spec);
        byte[] encData = hmacSHA256.doFinal(content.getBytes(UTF_8));
        return byteArrayToHexString(encData);
    }

    private static String byteArrayToHexString(byte[] data)
    {
        StringBuilder sb = new StringBuilder();
        String s;
        for(int n = 0; data != null && n < data.length; n++)
        {
            s = Integer.toHexString(data[n] & 0XFF);
            if(s.length() == 1)
            {
                sb.append('0');
            }
            sb.append(s);
        }
        return sb.toString().toLowerCase();
    }

    public static void main(String[] args)
            throws NoSuchAlgorithmException, InvalidKeyException, UnsupportedEncodingException
    {
        Map<String, Object> m = new HashMap<>();
        m.put("timestamp", 1525683727);
        m.put("idNo", "110102199001012279");
        m.put("name", "test");
        m.put("phone", "13624793586");
        SortedMap sortedMap = MapUtils.sortKey(m);
        String content = JSON.toJSONString(sortedMap);
        System.out.println("content: " + content);
        System.out.println("sign: " + sign(content, "61c072ac63db41e0875657edc2533a94"));
    }
}
