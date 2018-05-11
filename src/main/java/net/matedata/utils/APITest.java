package net.matedata.utils;

import com.alibaba.fastjson.JSON;
import com.alibaba.fastjson.JSONArray;
import com.alibaba.fastjson.JSONObject;
import net.matedata.model.HologramResp;
import net.matedata.model.Person;
import org.apache.http.HttpEntity;
import org.apache.http.util.EntityUtils;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileOutputStream;
import java.io.IOException;
import java.io.OutputStream;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;
import java.util.SortedMap;

public class APITest
{
    private static final String URL_QXBG = "https://api.matedata.net/credit/hologram";
    private static final String INPUT_FILE = "e://to_test/%s.xlsx";
    private static final String OUTPUT_FILE = "e://to_test/%s-全息报告.xlsx";
    private static final String TEMPLATE =
            "{\"timestamp\":%d,\"sign\":\"%s\",\"payload\":{\"phone\":\"%s\",\"name\":\"%s\",\"idNo\":\"%s\"}}";

    private static String secretKey; // 用户密钥
    private static String accessToken; // 用户访问凭证
    private static Map<String, String> headers; // 请求头

    public static void main(String[] args) throws IOException
    {
        String filename = "20180511";  // TODO 带测试的文件名
        secretKey = "757c9dd74b764fb9bd94302b13fef8af"; // TODO 用户密钥
        accessToken = "lfkc95k36eo5t8oche29z1h4y4z7d6v4o988y3j9p4i9u30a4cb2peg5q84b44qa"; // TODO 用户访问凭证
        // 构建请求头信息
        buildRequestHeader();
        String inputFile = String.format(INPUT_FILE, filename);
        String outputFile = String.format(OUTPUT_FILE, filename);
        // 读取用户信息列表
        List<Person> personList = readPersonInfoFromExcel(inputFile);
        System.out.println("总计" + personList.size() + "个用户待查询...");
        List<HologramResp> respList = new ArrayList<>(personList.size());
        HologramResp resp;
        int i = 0;
        for(Person person : personList)
        {
            resp = getPersonHologramInfo(person);
            if(resp != null)
            {
                respList.add(resp);
            }
            i++;
            if(i % 100 == 0)
            {
                System.out.println("当前已查询" + i);
            }
        }
        writeOutputFile(outputFile, respList);
        System.out.println("测试结果输出成功：" + outputFile);
    }

    private static void buildRequestHeader()
    {
        headers = new HashMap<>();
        headers.put("accept-version", "1.0");
        headers.put("Content-Type", "application/json");
        headers.put("Access-Token", accessToken);
    }

    private static List<Person> readPersonInfoFromExcel(String excelFile) throws IOException
    {
        List<Person> list = new ArrayList<>();
        Workbook wb = new XSSFWorkbook(excelFile);
        // 自定义id，姓名，身份证，手机号
        Sheet sheet = wb.getSheetAt(0);
        System.out.println(excelFile + ": 总计" + sheet.getPhysicalNumberOfRows() + "行");
        Person p;
        int i = 0;
        for(Row row : sheet)
        {
            i++;
            // 仅以文本格式读取，如报错，请将文件内容格式转换后重试
            String id, name, idNo, phone;
            try
            {
                id = row.getCell(0).getStringCellValue();
                name = row.getCell(1).getStringCellValue();
                idNo = row.getCell(2).getStringCellValue();
                phone = row.getCell(3).getStringCellValue();
            }
            catch(Exception e)
            {
                System.out.println("read error in line " + i);
                continue;
            }
            p = new Person();
            p.setId(id);
            p.setName(name);
            p.setIdNo(idNo);
            p.setPhone(phone);
            list.add(p);
        }
        System.out.println("已读入" + list.size() + "个用户信息");
        return list;
    }

    private static HologramResp getPersonHologramInfo(Person person)
    {
        HologramResp hr = new HologramResp();
        hr.setPerson(person);
        try
        {
            long ts = System.currentTimeMillis() / 1000;
            JSONObject p = (JSONObject) JSON.toJSON(person);
            p.remove("id");
            p.put("timestamp", ts);
            SortedMap sortedMap = MapUtils.sortKey(p);
            String sign = SignUtils.sign(JSON.toJSONString(sortedMap), secretKey);
            String content = String.format(TEMPLATE, ts, sign, person.getPhone(), person.getName(), person.getIdNo());
            HttpEntity entity = HttpUtils.postData(URL_QXBG, headers, content);
            String s = EntityUtils.toString(entity);
            JSONObject json = JSON.parseObject(s);
            // System.out.println(person.getPhone() + ": " + s);
            boolean hit = json.getInteger("code") == 0 && json.getInteger("resCode") == 0;
            if(hit)
            {
                hr.setHit(1);
                JSONObject res = json.getJSONObject("resData");
                hr.setBlacklistCode(res.getInteger("blacklistCode"));
                hr.setOverdueCode(res.getInteger("overdueCode"));
                hr.setOverdueRes(res.getJSONObject("overdueRes"));
                hr.setMultiRegisterCode(res.getInteger("multiRegisterCode"));
                hr.setMultiRegisterRes(res.getJSONObject("multiRegisterRes"));
                hr.setLoanCode(res.getInteger("creditableCode"));
                hr.setLoanRes(res.getJSONObject("creditableRes"));
                hr.setRefuseCode(res.getInteger("rejecteeCode"));
                hr.setRefuseRes(res.getJSONObject("rejecteeRes"));
                hr.setRepaymentCode(res.getInteger("favourableCode"));
                hr.setRepaymentRes(res.getJSONObject("favourableRes"));
            }
            else
            {
                hr.setHit(0);
                hr.setBlacklistCode(1);
                hr.setOverdueCode(1);
                hr.setMultiRegisterCode(1);
                hr.setLoanCode(1);
                hr.setRefuseCode(1);
                hr.setRepaymentCode(1);
            }
        }
        catch(Exception e)
        {
            System.out.println("error: " + person.getPhone());
            return null;
        }
        return hr;
    }

    private static void writeOutputFile(String outputFile, List<HologramResp> respList) throws IOException
    {
        Workbook wb = new XSSFWorkbook();
        Sheet sheet = wb.createSheet("sheet");
        Row row = sheet.createRow(0);
        row.createCell(0).setCellValue("序号");
        row.createCell(1).setCellValue("编号");
        row.createCell(2).setCellValue("姓名");
        row.createCell(3).setCellValue("身份证");
        row.createCell(4).setCellValue("手机");
        row.createCell(5).setCellValue("查询状态");
        row.createCell(6).setCellValue("1. 黑名单");
        row.createCell(7).setCellValue("2. 常欠客");
        row.createCell(8).setCellValue("3. 多头客");
        row.createCell(9).setCellValue("4. 拒贷客");
        row.createCell(10).setCellValue("5. 通过客");
        row.createCell(11).setCellValue("6. 优良客");
        if(respList != null && respList.size() > 0)
        {
            int i = 1;
            for(HologramResp hr : respList)
            {
                Row _row = sheet.createRow(i);
                _row.createCell(0).setCellValue(i);
                _row.createCell(1).setCellValue(hr.getPerson().getId());
                _row.createCell(2).setCellValue(hr.getPerson().getName());
                _row.createCell(3).setCellValue(hr.getPerson().getIdNo());
                _row.createCell(4).setCellValue(hr.getPerson().getPhone());
                _row.createCell(5).setCellValue(hr.getHit());
                _row.createCell(6).setCellValue(hr.getBlacklistCode() == 1 ? "未命中" : "命中");
                _row.createCell(7).setCellValue(hr.getOverdueCode() == 1 ? "未命中" : getOverdueDesc(hr.getOverdueRes()));
                _row.createCell(8).setCellValue(
                        hr.getMultiRegisterCode() == 1 ? "未命中" : getMultiRegisterDesc(hr.getMultiRegisterRes()));
                _row.createCell(9).setCellValue(hr.getRefuseCode() == 1 ? "未命中" : getRefuseDesc(hr.getRefuseRes()));
                _row.createCell(10).setCellValue(hr.getLoanCode() == 1 ? "未命中" : getLoanDesc(hr.getLoanRes()));
                _row.createCell(11).setCellValue(
                        hr.getRepaymentCode() == 1 ? "未命中" : getRepaymentDesc(hr.getRepaymentRes()));
                i++;
            }
        }
        OutputStream out = new FileOutputStream(outputFile);
        wb.write(out);
        out.close();
        wb.close();
    }

    private static String getOverdueDesc(JSONObject obj)
    {
        if(obj == null)
        {
            return "";
        }
        StringBuilder sb = new StringBuilder();
        sb.append("历史最大逾期金额：").append(obj.getString("overdueAmountMax")).append("\n");
        sb.append("历史最长逾期天数：").append(obj.getString("overdueDaysMax")).append("\n");
        sb.append("逾期最早出现时间：").append(obj.getString("overdueEarliestTime")).append("\n");
        sb.append("逾期最近出现时间：").append(obj.getString("overdueLatestTime")).append("\n");
        sb.append("历史逾期平台总数：").append(obj.getString("overduePlatformTotal")).append("\n");
        sb.append("今日逾期平台个数：").append(obj.getString("overduePlatformToday")).append("\n");
        sb.append("近3天逾期平台个数：").append(obj.getString("overduePlatform3Days")).append("\n");
        sb.append("近7天逾期平台个数：").append(obj.getString("overduePlatform7Days")).append("\n");
        sb.append("近15天逾期平台个数：").append(obj.getString("overduePlatform15Days")).append("\n");
        sb.append("近30天逾期平台个数：").append(obj.getString("overduePlatform30Days")).append("\n");
        sb.append("近60天逾期平台个数：").append(obj.getString("overduePlatform60Days")).append("\n");
        sb.append("近90天逾期平台个数：").append(obj.getString("overduePlatform90Days")).append("\n");
        JSONArray arr = obj.getJSONArray("overduePlatforms");
        JSONObject pt;
        if(arr != null && arr.size() > 0)
        {
            sb.append("逾期平台：\n");
            for(int i = 0; i < arr.size(); i++)
            {
                pt = arr.getJSONObject(i);
                sb.append(i + 1)
                        .append(". 平台编号：")
                        .append(pt.getString("platformCode"))
                        .append("，平台类型：")
                        .append(pt.getString("platformType"))
                        .append("，最早逾期时间：")
                        .append(pt.getString("overdueEarliestTime"))
                        .append("，最近逾期时间：")
                        .append(pt.getString("overdueLatestTime"))
                        .append("\n");
            }
        }
        return sb.toString();
    }

    private static String getMultiRegisterDesc(JSONObject obj)
    {
        if(obj == null)
        {
            return "";
        }
        StringBuilder sb = new StringBuilder();
        sb.append("历史单日最多注册平台个数：").append(obj.getString("registerInOneDayMax")).append("\n");
        sb.append("注册最早出现时间：").append(obj.getString("registerEarliestDate")).append("\n");
        sb.append("注册最近出现时间：").append(obj.getString("registerLatestDate")).append("\n");
        sb.append("历史注册平台总数：").append(obj.getString("registerPlatformTotal")).append("\n");
        sb.append("今日注册平台个数：").append(obj.getString("registerPlatformToday")).append("\n");
        sb.append("近3天注册平台个数：").append(obj.getString("registerPlatform3Days")).append("\n");
        sb.append("近7天注册平台个数：").append(obj.getString("registerPlatform7Days")).append("\n");
        sb.append("近15天注册平台个数：").append(obj.getString("registerPlatform15Days")).append("\n");
        sb.append("近30天注册平台个数：").append(obj.getString("registerPlatform30Days")).append("\n");
        sb.append("近60天注册平台个数：").append(obj.getString("registerPlatform60Days")).append("\n");
        sb.append("近90天注册平台个数：").append(obj.getString("registerPlatform90Days")).append("\n");
        JSONArray arr = obj.getJSONArray("registerPlatforms");
        JSONObject pt;
        if(arr != null && arr.size() > 0)
        {
            sb.append("注册平台：\n");
            for(int i = 0; i < arr.size(); i++)
            {
                pt = arr.getJSONObject(i);
                sb.append(i + 1).append(". 平台编号：").append(pt.getString("platformCode")).append("，平台类型：").append(
                        pt.getString("platformType")).append("，注册时间：").append(pt.getString("registerDate")).append(
                        "\n");
            }
        }
        return sb.toString();
    }

    private static String getRefuseDesc(JSONObject obj)
    {
        if(obj == null)
        {
            return "";
        }
        StringBuilder sb = new StringBuilder();
        sb.append("历史最大拒贷金额：").append(obj.getString("refuseAmountMax")).append("\n");
        sb.append("拒贷最早出现时间：").append(obj.getString("refuseEarliestDate")).append("\n");
        sb.append("拒贷最近出现时间：").append(obj.getString("refuseLatestDate")).append("\n");
        sb.append("历史拒贷平台总数：").append(obj.getString("refusePlatformTotal")).append("\n");
        sb.append("今日拒贷平台个数：").append(obj.getString("refusePlatformToday")).append("\n");
        sb.append("近3天拒贷平台个数：").append(obj.getString("refusePlatform3Days")).append("\n");
        sb.append("近7天拒贷平台个数：").append(obj.getString("refusePlatform7Days")).append("\n");
        sb.append("近15天拒贷平台个数：").append(obj.getString("refusePlatform15Days")).append("\n");
        sb.append("近30天拒贷平台个数：").append(obj.getString("refusePlatform30Days")).append("\n");
        sb.append("近60天拒贷平台个数：").append(obj.getString("refusePlatform60Days")).append("\n");
        sb.append("近90天拒贷平台个数：").append(obj.getString("refusePlatform90Days")).append("\n");
        JSONArray arr = obj.getJSONArray("refusePlatforms");
        JSONObject pt;
        if(arr != null && arr.size() > 0)
        {
            sb.append("拒贷平台：\n");
            for(int i = 0; i < arr.size(); i++)
            {
                pt = arr.getJSONObject(i);
                sb.append(i + 1)
                        .append(". 平台编号：")
                        .append(pt.getString("platformCode"))
                        .append("，平台类型：")
                        .append(pt.getString("platformType"))
                        .append("，最早拒贷时间：")
                        .append(pt.getString("refuseEarliestDate"))
                        .append("，最近拒贷时间：")
                        .append(pt.getString("refuseLatestDate"))
                        .append("\n");
            }
        }
        return sb.toString();
    }

    private static String getLoanDesc(JSONObject obj)
    {
        if(obj == null)
        {
            return "";
        }
        StringBuilder sb = new StringBuilder();
        sb.append("历史最大放款金额：").append(obj.getString("loanAmountMax")).append("\n");
        sb.append("放款最早出现时间：").append(obj.getString("loanEarliestDate")).append("\n");
        sb.append("放款最近出现时间：").append(obj.getString("loanLatestDate")).append("\n");
        sb.append("历史放款平台总数：").append(obj.getString("loanPlatformTotal")).append("\n");
        sb.append("今日放款平台个数：").append(obj.getString("loanPlatformToday")).append("\n");
        sb.append("近3天放款平台个数：").append(obj.getString("loanPlatform3Days")).append("\n");
        sb.append("近7天放款平台个数：").append(obj.getString("loanPlatform7Days")).append("\n");
        sb.append("近15天放款平台个数：").append(obj.getString("loanPlatform15Days")).append("\n");
        sb.append("近30天放款平台个数：").append(obj.getString("loanPlatform30Days")).append("\n");
        sb.append("近60天放款平台个数：").append(obj.getString("loanPlatform60Days")).append("\n");
        sb.append("近90天放款平台个数：").append(obj.getString("loanPlatform90Days")).append("\n");
        JSONArray arr = obj.getJSONArray("loanPlatforms");
        JSONObject pt;
        if(arr != null && arr.size() > 0)
        {
            sb.append("放款平台：\n");
            for(int i = 0; i < arr.size(); i++)
            {
                pt = arr.getJSONObject(i);
                sb.append(i + 1)
                        .append(". 平台编号：")
                        .append(pt.getString("platformCode"))
                        .append("，平台类型：")
                        .append(pt.getString("platformType"))
                        .append("，最早放款时间：")
                        .append(pt.getString("loanEarliestDate"))
                        .append("，最近放款时间：")
                        .append(pt.getString("loanLatestDate"))
                        .append("\n");
            }
        }
        return sb.toString();
    }

    private static String getRepaymentDesc(JSONObject obj)
    {
        if(obj == null)
        {
            return "";
        }
        StringBuilder sb = new StringBuilder();
        sb.append("历史最大还款金额：").append(obj.getString("repaymentAmountMax")).append("\n");
        sb.append("还款最早出现时间：").append(obj.getString("repaymentEarliestDate")).append("\n");
        sb.append("还款最近出现时间：").append(obj.getString("repaymentLatestDate")).append("\n");
        sb.append("历史还款平台总数：").append(obj.getString("repaymentPlatformTotal")).append("\n");
        sb.append("今日还款平台个数：").append(obj.getString("repaymentPlatformToday")).append("\n");
        sb.append("近3天还款平台个数：").append(obj.getString("repaymentPlatform3Days")).append("\n");
        sb.append("近7天还款平台个数：").append(obj.getString("repaymentPlatform7Days")).append("\n");
        sb.append("近15天还款平台个数：").append(obj.getString("repaymentPlatform15Days")).append("\n");
        sb.append("近30天还款平台个数：").append(obj.getString("repaymentPlatform30Days")).append("\n");
        sb.append("近60天还款平台个数：").append(obj.getString("repaymentPlatform60Days")).append("\n");
        sb.append("近90天还款平台个数：").append(obj.getString("repaymentPlatform90Days")).append("\n");
        JSONArray arr = obj.getJSONArray("repaymentPlatforms");
        JSONObject pt;
        if(arr != null && arr.size() > 0)
        {
            sb.append("还款平台：\n");
            for(int i = 0; i < arr.size(); i++)
            {
                pt = arr.getJSONObject(i);
                sb.append(i + 1)
                        .append(". 平台编号：")
                        .append(pt.getString("platformCode"))
                        .append("，平台类型：")
                        .append(pt.getString("platformType"))
                        .append("，最早还款时间：")
                        .append(pt.getString("repaymentEarliestDate"))
                        .append("，最近还款时间：")
                        .append(pt.getString("repaymentLatestDate"))
                        .append("\n");
            }
        }
        return sb.toString();
    }
}
