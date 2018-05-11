package net.matedata.utils;

import org.apache.http.HttpEntity;
import org.apache.http.HttpResponse;
import org.apache.http.client.HttpClient;
import org.apache.http.client.methods.HttpPost;
import org.apache.http.entity.StringEntity;
import org.apache.http.impl.client.HttpClients;

import java.io.IOException;
import java.util.Map;

public class HttpUtils
{
    private static final String DEFAULT_CHARSET = "UTF-8";

    private HttpUtils()
    {
    }

    public static HttpEntity postData(String uri, Map<String, String> headers, String content) throws IOException
    {
        HttpClient client = HttpClients.createDefault();
        HttpPost post = new HttpPost(uri);
        post.setEntity(new StringEntity(content, DEFAULT_CHARSET));
        if(headers != null)
        {
            for(Map.Entry<String, String> entry : headers.entrySet())
            {
                post.addHeader(entry.getKey(), entry.getValue());
            }
        }
        HttpResponse response = client.execute(post);
        return response.getEntity();
    }
}
