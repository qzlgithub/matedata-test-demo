package net.matedata.utils;

import com.alibaba.fastjson.JSON;
import com.alibaba.fastjson.JSONArray;
import com.alibaba.fastjson.JSONObject;
import net.matedata.model.HologramResp;
import net.matedata.model.Person;
import org.apache.http.HttpEntity;
import org.apache.http.util.EntityUtils;
import org.apache.poi.ss.usermodel.BorderStyle;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.FillPatternType;
import org.apache.poi.ss.usermodel.HorizontalAlignment;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.VerticalAlignment;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.XSSFFont;
import org.apache.poi.xssf.usermodel.XSSFFormulaEvaluator;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileOutputStream;
import java.io.IOException;
import java.io.OutputStream;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.HashSet;
import java.util.List;
import java.util.Map;
import java.util.Set;

public class APITest
{
    private static final String URL_QXBG =
            "https://api.matedata.net/detection/hologram/HLGep8XB2ItKqElp1ktuh5SpDYzLQ8xz";
    private static final String INPUT_FILE = "C:\\Users\\JinKX\\Desktop\\to_test\\test\\%s.xlsx";
    private static final String OUTPUT_FILE = "C:\\Users\\JinKX\\Desktop\\to_test\\test\\%s-全息报告.xlsx";
    static int c;

    private static Map<String, String> headers; // 请求头

    public static void main(String[] args) throws IOException
    {
        String filename = "Test";  // TODO 带测试的文件名
        // 构建请求头信息
        buildRequestHeader();
        String inputFile = String.format(INPUT_FILE, filename);
        String outputFile = String.format(OUTPUT_FILE, filename);
        // 读取用户信息列表
        List<Person> personList = readPersonInfoFromExcel(inputFile);
        System.out.println("总计" + personList.size() + "个用户待查询...");
        if(personList.size() == 0)
        {
            return;
        }
        List<HologramResp> respList = new ArrayList<>(personList.size());
        List<Person> _personList = new ArrayList<>(100);
        List<HologramResp> resList;
        for(Person p : personList)
        {
            _personList.add(p);
            if(_personList.size() == 100)
            {
                resList = getPersonHologramInfo(_personList);
                respList.addAll(resList);
                System.out.println("100个已查询");
                _personList = new ArrayList<>(100);
            }
        }
        if(_personList.size() != 0)
        {
            resList = getPersonHologramInfo(_personList);
            respList.addAll(resList);
            System.out.println(_personList.size() + "个已查询");
        }

        writeOutputFile(outputFile, respList);
        System.out.println("测试结果输出成功：" + outputFile);
    }

    private static void buildRequestHeader()
    {
        headers = new HashMap<>();
        headers.put("Content-Type", "application/json");
    }

    private static List<Person> readPersonInfoFromExcel(String excelFile) throws IOException
    {
        List<Person> list = new ArrayList<>();
        Set<String> phoneList = new HashSet<>();
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
            p.setId(id.trim());
            p.setName(name.trim());
            p.setIdNo(idNo.trim());
            p.setPhone(phone.trim());
            if(phone.trim().length() == 11)
            {
                if(!phoneList.contains(phone.trim()))
                {
                    list.add(p);
                    phoneList.add(phone.trim());
                }
            }
        }
        System.out.println("已读入" + list.size() + "个用户信息");
        return list;
    }

    private static List<HologramResp> getPersonHologramInfo(List<Person> personList) throws IOException
    {
        List<HologramResp> list = new ArrayList<>(personList.size());
        try
        {
            List<String> phoneList = new ArrayList<>(personList.size());
            for(Person person : personList)
            {
                phoneList.add(person.getPhone());
            }
            JSONObject content = new JSONObject();
            content.put("phones", phoneList);
            HttpEntity entity = HttpUtils.postData(URL_QXBG, headers, JSON.toJSONString(content));
            String s = EntityUtils.toString(entity);
            JSONObject json = JSON.parseObject(s);
            JSONObject resData = json.getJSONObject("resData");
            HologramResp hr;
            for(Person person : personList)
            {
                hr = new HologramResp();
                hr.setPerson(person);
                JSONObject res = resData.getJSONObject(person.getPhone());
                if(res != null && (res.getInteger("blacklistCode") == 0 || res.getInteger("overdueCode") == 0 ||
                        res.getInteger("multiRegisterCode") == 0 || res.getInteger("creditableCode") == 0 ||
                        res.getInteger("rejecteeCode") == 0 || res.getInteger("favourableCode") == 0))
                {
                    hr.setHit(1);
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
                list.add(hr);
            }
        }
        catch(Exception e)
        {
            System.out.println("error: " + e.getMessage());
        }
        return list;
    }

    private static void writeOutputFile(String outputFile, List<HologramResp> respList) throws IOException
    {
        XSSFWorkbook wb = new XSSFWorkbook();
        Sheet sheet = wb.createSheet("sheet");
        //冻结窗口
        sheet.createFreezePane(4, 2);

        Row row = sheet.createRow(0);
        Cell cell2 = row.createCell(0);
        cell2.setCellValue("姓名");
        Cell cell3 = row.createCell(1);
        cell3.setCellValue("身份证");
        Cell cell4 = row.createCell(2);
        cell4.setCellValue("手机");
        Cell cell5 = row.createCell(3);
        cell5.setCellValue("查询状态\n（1-命中；0-未命中）");
        Cell cell6 = row.createCell(4);
        cell6.setCellValue("1. 黑名单");
        Cell cell7 = row.createCell(5);
        cell7.setCellValue("2. 常欠客");
        Cell cell8 = row.createCell(19);
        cell8.setCellValue("3. 多头客");
        Cell cell9 = row.createCell(32);
        cell9.setCellValue("4. 拒贷客");
        Cell cell10 = row.createCell(45);
        cell10.setCellValue("5. 通过客");
        Cell cell11 = row.createCell(58);
        cell11.setCellValue("6. 优良客");
        Cell cell12 = row.createCell(70);
        cell12.setCellValue("");

        Row row1 = sheet.createRow(1);
        Cell cel0 = row1.createCell(4);
        cel0.setCellValue("是否命中");
        Cell cel = row1.createCell(5);
        cel.setCellValue("是否命中");
        Cell cel1 = row1.createCell(6);
        cel1.setCellValue("历史最大逾期金额");
        Cell cel2 = row1.createCell(7);
        cel2.setCellValue("历史最长逾期天数");
        Cell cel3 = row1.createCell(8);
        cel3.setCellValue("逾期最早出现时间");
        Cell cel4 = row1.createCell(9);
        cel4.setCellValue("逾期最近出现时间");
        Cell cel5 = row1.createCell(10);
        cel5.setCellValue("历史逾期平台总数");
        Cell cel6 = row1.createCell(11);
        cel6.setCellValue("今日逾期平台个数");
        Cell cel7 = row1.createCell(12);
        cel7.setCellValue("近3天逾期平台个数");
        Cell cel8 = row1.createCell(13);
        cel8.setCellValue("近7天逾期平台个数");
        Cell cel9 = row1.createCell(14);
        cel9.setCellValue("近15天逾期平台个数");
        Cell cel10 = row1.createCell(15);
        cel10.setCellValue("近30天逾期平台个数");
        Cell cel11 = row1.createCell(16);
        cel11.setCellValue("近60天逾期平台个数");
        Cell cel12 = row1.createCell(17);
        cel12.setCellValue("近90天逾期平台个数");
        Cell cel13 = row1.createCell(18);
        cel13.setCellValue("逾期平台");
        Cell cel14 = row1.createCell(19);
        cel14.setCellValue("是否命中");
        Cell cel15 = row1.createCell(20);
        cel15.setCellValue("历史单日最多注册平台个数");
        Cell cel16 = row1.createCell(21);
        cel16.setCellValue("注册最早出现时间");
        Cell cel17 = row1.createCell(22);
        cel17.setCellValue("注册最近出现时间");
        Cell cel18 = row1.createCell(23);
        cel18.setCellValue("历史注册平台总数");
        Cell cel19 = row1.createCell(24);
        cel19.setCellValue("今日注册平台个数");
        Cell cel20 = row1.createCell(25);
        cel20.setCellValue("近3天注册平台个数");
        Cell cel21 = row1.createCell(26);
        cel21.setCellValue("近7天注册平台个数");
        Cell cel22 = row1.createCell(27);
        cel22.setCellValue("近15天注册平台个数");
        Cell cel23 = row1.createCell(28);
        cel23.setCellValue("近30天注册平台个数");
        Cell cel24 = row1.createCell(29);
        cel24.setCellValue("近60天注册平台个数");
        Cell cel25 = row1.createCell(30);
        cel25.setCellValue("近90天注册平台个数");
        Cell cel26 = row1.createCell(31);
        cel26.setCellValue("注册平台");
        Cell cel27 = row1.createCell(32);
        cel27.setCellValue("是否命中");
        Cell cel28 = row1.createCell(33);
        cel28.setCellValue("历史最大拒贷金额");
        Cell cel29 = row1.createCell(34);
        cel29.setCellValue("拒贷最早出现时间");
        Cell cel30 = row1.createCell(35);
        cel30.setCellValue("拒贷最近出现时间");
        Cell cel31 = row1.createCell(36);
        cel31.setCellValue("历史拒贷平台总数");
        Cell cel32 = row1.createCell(37);
        cel32.setCellValue("今日拒贷平台个数");
        Cell cel33 = row1.createCell(38);
        cel33.setCellValue("近3天拒贷平台个数");
        Cell cel34 = row1.createCell(39);
        cel34.setCellValue("近7天拒贷平台个数");
        Cell cel35 = row1.createCell(40);
        cel35.setCellValue("近15天拒贷平台个数");
        Cell cel36 = row1.createCell(41);
        cel36.setCellValue("近30天拒贷平台个数");
        Cell cel37 = row1.createCell(42);
        cel37.setCellValue("近60天拒贷平台个数");
        Cell cel38 = row1.createCell(43);
        cel38.setCellValue("近90天拒贷平台个数");
        Cell cel39 = row1.createCell(44);
        cel39.setCellValue("拒贷平台");
        Cell cel40 = row1.createCell(45);
        cel40.setCellValue("是否命中");
        Cell cel41 = row1.createCell(46);
        cel41.setCellValue("历史最大放款金额");
        Cell cel42 = row1.createCell(47);
        cel42.setCellValue("放款最早出现时间");
        Cell cel43 = row1.createCell(48);
        cel43.setCellValue("放款最近出现时间");
        Cell cel44 = row1.createCell(49);
        cel44.setCellValue("历史放款平台总数");
        Cell cel45 = row1.createCell(50);
        cel45.setCellValue("今日放款平台个数");
        Cell cel46 = row1.createCell(51);
        cel46.setCellValue("近3天放款平台个数");
        Cell cel47 = row1.createCell(52);
        cel47.setCellValue("近7天放款平台个数");
        Cell cel48 = row1.createCell(53);
        cel48.setCellValue("近15天放款平台个数");
        Cell cel49 = row1.createCell(54);
        cel49.setCellValue("近30天放款平台个数");
        Cell cel50 = row1.createCell(55);
        cel50.setCellValue("近60天放款平台个数");
        Cell cel51 = row1.createCell(56);
        cel51.setCellValue("近90天放款平台个数");
        Cell cel52 = row1.createCell(57);
        cel52.setCellValue("放款平台");
        Cell cel53 = row1.createCell(58);
        cel53.setCellValue("是否命中");
        Cell cel54 = row1.createCell(59);
        cel54.setCellValue("历史最大还款金额");
        Cell cel55 = row1.createCell(60);
        cel55.setCellValue("还款最早出现时间");
        Cell cel56 = row1.createCell(61);
        cel56.setCellValue("还款最近出现时间");
        Cell cel57 = row1.createCell(62);
        cel57.setCellValue("历史还款平台总数");
        Cell cel58 = row1.createCell(63);
        cel58.setCellValue("今日还款平台个数");
        Cell cel59 = row1.createCell(64);
        cel59.setCellValue("近3天还款平台个数");
        Cell cel60 = row1.createCell(65);
        cel60.setCellValue("近7天还款平台个数");
        Cell cel61 = row1.createCell(66);
        cel61.setCellValue("近15天还款平台个数");
        Cell cel62 = row1.createCell(67);
        cel62.setCellValue("近30天还款平台个数");
        Cell cel63 = row1.createCell(68);
        cel63.setCellValue("近60天还款平台个数");
        Cell cel64 = row1.createCell(69);
        cel64.setCellValue("近90天还款平台个数");
        Cell cel65 = row1.createCell(70);
        cel65.setCellValue("还款平台");

        //合并单元格
        sheet.addMergedRegion(new CellRangeAddress(0, 1, 0, 0));
        sheet.addMergedRegion(new CellRangeAddress(0, 1, 1, 1));
        sheet.addMergedRegion(new CellRangeAddress(0, 1, 2, 2));
        sheet.addMergedRegion(new CellRangeAddress(0, 1, 3, 3));
        sheet.addMergedRegion(new CellRangeAddress(0, 0, 5, 18));
        sheet.addMergedRegion(new CellRangeAddress(0, 0, 19, 31));
        sheet.addMergedRegion(new CellRangeAddress(0, 0, 32, 44));
        sheet.addMergedRegion(new CellRangeAddress(0, 0, 45, 57));
        sheet.addMergedRegion(new CellRangeAddress(0, 0, 58, 70));

        //行高
        //row.setHeightInPoints(35);

        CellStyle style0 = wb.createCellStyle();
        //居中
        style0.setAlignment(HorizontalAlignment.CENTER);
        style0.setVerticalAlignment(VerticalAlignment.CENTER);
        //自动换行
        style0.setWrapText(true);

        cell2.setCellStyle(style0);
        cell3.setCellStyle(style0);
        cell4.setCellStyle(style0);
        cell5.setCellStyle(style0);

        //第二行文字的自动换行和居中设置
        cel1.setCellStyle(style0);
        cel2.setCellStyle(style0);
        cel3.setCellStyle(style0);
        cel4.setCellStyle(style0);
        cel5.setCellStyle(style0);
        cel6.setCellStyle(style0);
        cel7.setCellStyle(style0);
        cel8.setCellStyle(style0);
        cel9.setCellStyle(style0);
        cel10.setCellStyle(style0);
        cel11.setCellStyle(style0);
        cel12.setCellStyle(style0);
        cel13.setCellStyle(style0);
        cel14.setCellStyle(style0);
        cel15.setCellStyle(style0);
        cel16.setCellStyle(style0);
        cel17.setCellStyle(style0);
        cel18.setCellStyle(style0);
        cel19.setCellStyle(style0);
        cel20.setCellStyle(style0);
        cel21.setCellStyle(style0);
        cel22.setCellStyle(style0);
        cel23.setCellStyle(style0);
        cel24.setCellStyle(style0);
        cel25.setCellStyle(style0);
        cel26.setCellStyle(style0);
        cel27.setCellStyle(style0);
        cel28.setCellStyle(style0);
        cel29.setCellStyle(style0);
        cel30.setCellStyle(style0);
        cel31.setCellStyle(style0);
        cel32.setCellStyle(style0);
        cel33.setCellStyle(style0);
        cel34.setCellStyle(style0);
        cel35.setCellStyle(style0);
        cel36.setCellStyle(style0);
        cel37.setCellStyle(style0);
        cel38.setCellStyle(style0);
        cel39.setCellStyle(style0);
        cel40.setCellStyle(style0);
        cel41.setCellStyle(style0);
        cel42.setCellStyle(style0);
        cel43.setCellStyle(style0);
        cel18.setCellStyle(style0);
        cel19.setCellStyle(style0);
        cel20.setCellStyle(style0);
        cel21.setCellStyle(style0);
        cel22.setCellStyle(style0);
        cel23.setCellStyle(style0);
        cel24.setCellStyle(style0);
        cel25.setCellStyle(style0);
        cel26.setCellStyle(style0);
        cel27.setCellStyle(style0);
        cel28.setCellStyle(style0);
        cel29.setCellStyle(style0);
        cel30.setCellStyle(style0);
        cel31.setCellStyle(style0);
        cel32.setCellStyle(style0);
        cel33.setCellStyle(style0);
        cel34.setCellStyle(style0);
        cel35.setCellStyle(style0);
        cel36.setCellStyle(style0);
        cel37.setCellStyle(style0);
        cel38.setCellStyle(style0);
        cel39.setCellStyle(style0);
        cel40.setCellStyle(style0);
        cel45.setCellStyle(style0);
        cel46.setCellStyle(style0);
        cel47.setCellStyle(style0);
        cel48.setCellStyle(style0);
        cel49.setCellStyle(style0);
        cel50.setCellStyle(style0);
        cel51.setCellStyle(style0);
        cel52.setCellStyle(style0);
        cel53.setCellStyle(style0);
        cel54.setCellStyle(style0);
        cel55.setCellStyle(style0);
        cel56.setCellStyle(style0);
        cel57.setCellStyle(style0);
        cel58.setCellStyle(style0);
        cel59.setCellStyle(style0);
        cel60.setCellStyle(style0);
        cel61.setCellStyle(style0);
        cel62.setCellStyle(style0);
        cel63.setCellStyle(style0);
        cel64.setCellStyle(style0);
        cel65.setCellStyle(style0);

        //前景填充色
        CellStyle style = wb.createCellStyle();
        style.setFillForegroundColor(IndexedColors.LIGHT_TURQUOISE.getIndex());
        style.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        //边框颜色设置
        style.setBorderBottom(BorderStyle.THIN);
        style.setBottomBorderColor(IndexedColors.GREY_25_PERCENT.getIndex());
        style.setBorderLeft(BorderStyle.THIN);
        style.setLeftBorderColor(IndexedColors.BLACK.getIndex());
        style.setBorderTop(BorderStyle.THIN);
        style.setTopBorderColor(IndexedColors.GREY_40_PERCENT.getIndex());
        //居中
        style.setAlignment(HorizontalAlignment.CENTER);
        style.setVerticalAlignment(VerticalAlignment.CENTER);
        cell6.setCellStyle(style);

        CellStyle style1 = wb.createCellStyle();
        style1.setFillForegroundColor(IndexedColors.CORNFLOWER_BLUE.getIndex());
        style1.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        //边框颜色设置
        style1.setBorderBottom(BorderStyle.THIN);
        style1.setBottomBorderColor(IndexedColors.GREY_25_PERCENT.getIndex());
        style1.setBorderLeft(BorderStyle.THIN);
        style1.setLeftBorderColor(IndexedColors.BLACK.getIndex());
        style1.setBorderTop(BorderStyle.THIN);
        style1.setTopBorderColor(IndexedColors.GREY_40_PERCENT.getIndex());
        //居中
        style1.setAlignment(HorizontalAlignment.CENTER);
        style1.setVerticalAlignment(VerticalAlignment.CENTER);
        cell7.setCellStyle(style1);

        CellStyle style2 = wb.createCellStyle();
        style2.setFillForegroundColor(IndexedColors.GREY_25_PERCENT.getIndex());
        style2.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        //边框
        style2.setBorderBottom(BorderStyle.THIN);
        style2.setBottomBorderColor(IndexedColors.GREY_40_PERCENT.getIndex());
        style2.setBorderLeft(BorderStyle.THIN);
        style2.setLeftBorderColor(IndexedColors.BLACK.getIndex());
        style2.setBorderRight(BorderStyle.THIN);
        style2.setRightBorderColor(IndexedColors.GREY_40_PERCENT.getIndex());
        /*style2.setBorderTop(BorderStyle.THIN);
        style2.setTopBorderColor(IndexedColors.GREY_40_PERCENT.getIndex());*/
        //居中
        style2.setAlignment(HorizontalAlignment.CENTER);
        style2.setVerticalAlignment(VerticalAlignment.CENTER);

        cell8.setCellStyle(style);
        cell9.setCellStyle(style1);
        cell10.setCellStyle(style);
        cell11.setCellStyle(style1);

        cel0.setCellStyle(style2);
        cel.setCellStyle(style2);
        cel14.setCellStyle(style2);
        cel27.setCellStyle(style2);
        cel40.setCellStyle(style2);
        cel53.setCellStyle(style2);

        if(respList != null && respList.size() > 0)
        {
            /*int c = 2;*/
            c = 2;
            // System.out.println(">>>: " + respList.size());
            for(HologramResp hr : respList)
            {
                // System.out.println(JSON.toJSONString(hr));
                JSONObject obj = hr.getOverdueRes();
                JSONObject obj2 = hr.getMultiRegisterRes();
                JSONObject obj3 = hr.getRefuseRes();
                JSONObject obj4 = hr.getLoanRes();
                JSONObject obj5 = hr.getRepaymentRes();

                Row _row = sheet.createRow(c);
                _row.createCell(0).setCellValue(hr.getPerson().getName());
                _row.createCell(1).setCellValue(hr.getPerson().getIdNo());
                _row.createCell(2).setCellValue(hr.getPerson().getPhone());
                _row.createCell(3).setCellValue(hr.getHit());

                Cell cel_row0 = _row.createCell(4);
                cel_row0.setCellValue(hr.getBlacklistCode() == 1 ? 0 : 1);
                cel_row0.setCellStyle(style2);
                /*_row.createCell(6).setCellValue(hr.getBlacklistCode() == 1 ? 0 : 1);*/

                //常欠客
                if(hr.getOverdueCode() == 1)
                {
                    Cell cel_row = _row.createCell(5);
                    cel_row.setCellValue(0);
                    cel_row.setCellStyle(style2);

                }
                else
                {
                    /*_row.createCell(7).setCellValue(1);*/
                    Cell cel_row = _row.createCell(5);
                    cel_row.setCellValue(1);
                    cel_row.setCellStyle(style2);
                    if(obj.getDouble("overdueAmountMax") != null)
                    {
                        _row.createCell(6).setCellValue(obj.getDouble("overdueAmountMax"));
                    }
                    if(obj.getInteger("overdueDaysMax") != null)
                    {
                        _row.createCell(7).setCellValue(obj.getInteger("overdueDaysMax"));
                    }
                    _row.createCell(8).setCellValue(obj.getString("overdueEarliestTime"));
                    _row.createCell(9).setCellValue(obj.getString("overdueLatestTime"));
                    if(obj.getInteger("overduePlatformTotal") != null)
                    {
                        _row.createCell(10).setCellValue(obj.getInteger("overduePlatformTotal"));

                    }
                    if(obj.getInteger("overduePlatformToday") != null)
                    {
                        _row.createCell(11).setCellValue(obj.getInteger("overduePlatformToday"));
                    }
                    if(obj.getInteger("overduePlatform3Days") != null)
                    {
                        _row.createCell(12).setCellValue(obj.getInteger("overduePlatform3Days"));

                    }
                    if(obj.getInteger("overduePlatform7Days") != null)
                    {
                        _row.createCell(13).setCellValue(obj.getInteger("overduePlatform7Days"));
                    }
                    if(obj.getInteger("overduePlatform15Days") != null)
                    {
                        _row.createCell(14).setCellValue(obj.getInteger("overduePlatform15Days"));
                    }
                    if(obj.getInteger("overduePlatform30Days") != null)
                    {
                        _row.createCell(15).setCellValue(obj.getInteger("overduePlatform30Days"));
                    }
                    if(obj.getInteger("overduePlatform60Days") != null)
                    {
                        _row.createCell(16).setCellValue(obj.getInteger("overduePlatform60Days"));
                    }
                    if(obj.getInteger("overduePlatform90Days") != null)
                    {
                        _row.createCell(17).setCellValue(obj.getInteger("overduePlatform90Days"));
                    }
                    JSONArray arr = obj.getJSONArray("overduePlatforms");
                    JSONObject pt;
                    StringBuilder sb = new StringBuilder();
                    if(arr != null && arr.size() > 0)
                    {
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
                    _row.createCell(18).setCellValue(sb.toString());
                }

                //多头客
                if(hr.getMultiRegisterCode() == 1)
                {
                    Cell cel_row = _row.createCell(19);
                    cel_row.setCellValue(0);
                    cel_row.setCellStyle(style2);
                }
                else
                {
                    Cell cel_row = _row.createCell(19);
                    cel_row.setCellValue(1);
                    cel_row.setCellStyle(style2);
                    if(obj2.getInteger("registerInOneDayMax") != null)
                    {
                        _row.createCell(20).setCellValue(obj2.getInteger("registerInOneDayMax"));
                    }
                    _row.createCell(21).setCellValue(obj2.getString("registerEarliestDate"));
                    _row.createCell(22).setCellValue(obj2.getString("registerLatestDate"));
                    if(obj2.getInteger("registerPlatformTotal") != null)
                    {
                        _row.createCell(23).setCellValue(obj2.getInteger("registerPlatformTotal"));
                    }
                    if(obj2.getInteger("registerPlatformToday") != null)
                    {
                        _row.createCell(24).setCellValue(obj2.getInteger("registerPlatformToday"));
                    }
                    if(obj2.getInteger("registerPlatform3Days") != null)
                    {
                        _row.createCell(25).setCellValue(obj2.getInteger("registerPlatform3Days"));
                    }
                    if(obj2.getInteger("registerPlatform7Days") != null)
                    {
                        _row.createCell(26).setCellValue(obj2.getInteger("registerPlatform7Days"));
                    }
                    if(obj2.getInteger("registerPlatform15Days") != null)
                    {
                        _row.createCell(27).setCellValue(obj2.getInteger("registerPlatform15Days"));
                    }
                    if(obj2.getInteger("registerPlatform30Days") != null)
                    {
                        _row.createCell(28).setCellValue(obj2.getInteger("registerPlatform30Days"));
                    }
                    if(obj2.getInteger("registerPlatform60Days") != null)
                    {
                        _row.createCell(29).setCellValue(obj2.getInteger("registerPlatform60Days"));
                    }
                    if(obj2.getInteger("registerPlatform90Days") != null)
                    {
                        _row.createCell(30).setCellValue(obj2.getInteger("registerPlatform90Days"));
                    }
                    JSONArray arr = obj2.getJSONArray("registerPlatforms");
                    JSONObject pt;
                    StringBuilder sb = new StringBuilder();
                    if(arr != null && arr.size() > 0)
                    {
                        for(int i = 0; i < arr.size(); i++)
                        {
                            pt = arr.getJSONObject(i);
                            sb.append(i + 1)
                                    .append(". 平台编号：")
                                    .append(pt.getString("platformCode"))
                                    .append("，平台类型：")
                                    .append(pt.getString("platformType"))
                                    .append("，注册时间：")
                                    .append(pt.getString("registerDate"))
                                    .append("\n");
                        }
                    }
                    _row.createCell(31).setCellValue(sb.toString());
                }

                //拒贷客
                if(hr.getRefuseCode() == 1)
                {
                    Cell cel_row = _row.createCell(32);
                    cel_row.setCellValue(0);
                    cel_row.setCellStyle(style2);
                }
                else
                {
                    Cell cel_row = _row.createCell(32);
                    cel_row.setCellValue(1);
                    cel_row.setCellStyle(style2);
                    if(obj3.getDouble("refuseAmountMax") != null)
                    {
                        _row.createCell(33).setCellValue(obj3.getDouble("refuseAmountMax"));
                    }
                    _row.createCell(34).setCellValue(obj3.getString("refuseEarliestDate"));
                    _row.createCell(35).setCellValue(obj3.getString("refuseLatestDate"));
                    if(obj3.getInteger("refusePlatformTotal") != null)
                    {
                        _row.createCell(36).setCellValue(obj3.getInteger("refusePlatformTotal"));
                    }
                    if(obj3.getInteger("refusePlatformToday") != null)
                    {
                        _row.createCell(37).setCellValue(obj3.getInteger("refusePlatformToday"));
                    }
                    if(obj3.getInteger("refusePlatform3Days") != null)
                    {
                        _row.createCell(38).setCellValue(obj3.getInteger("refusePlatform3Days"));
                    }
                    if(obj3.getInteger("refusePlatform7Days") != null)
                    {
                        _row.createCell(39).setCellValue(obj3.getInteger("refusePlatform7Days"));
                    }
                    if(obj3.getInteger("refusePlatform15Days") != null)
                    {
                        _row.createCell(40).setCellValue(obj3.getInteger("refusePlatform15Days"));
                    }
                    if(obj3.getInteger("refusePlatform30Days") != null)
                    {
                        _row.createCell(41).setCellValue(obj3.getInteger("refusePlatform30Days"));
                    }
                    if(obj3.getInteger("refusePlatform60Days") != null)
                    {
                        _row.createCell(42).setCellValue(obj3.getInteger("refusePlatform60Days"));

                    }
                    if(obj3.getInteger("refusePlatform90Days") != null)
                    {
                        _row.createCell(43).setCellValue(obj3.getInteger("refusePlatform90Days"));
                    }
                    JSONArray arr = obj3.getJSONArray("refusePlatforms");
                    JSONObject pt;
                    StringBuilder sb = new StringBuilder();
                    if(arr != null && arr.size() > 0)
                    {
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
                    _row.createCell(44).setCellValue(sb.toString());
                }

                //通过客
                if(hr.getLoanCode() == 1)
                {
                    Cell cel_row = _row.createCell(45);
                    cel_row.setCellValue(0);
                    cel_row.setCellStyle(style2);
                }
                else
                {
                    Cell cel_row = _row.createCell(45);
                    cel_row.setCellValue(1);
                    cel_row.setCellStyle(style2);
                    if(obj4.getDouble("loanAmountMax") != null)
                    {
                        _row.createCell(46).setCellValue(obj4.getDouble("loanAmountMax"));
                    }
                    _row.createCell(47).setCellValue(obj4.getString("loanEarliestDate"));
                    _row.createCell(48).setCellValue(obj4.getString("loanLatestDate"));
                    if(obj4.getInteger("loanPlatformTotal") != null)
                    {
                        _row.createCell(49).setCellValue(obj4.getInteger("loanPlatformTotal"));
                    }
                    if(obj4.getInteger("loanPlatformToday") != null)
                    {
                        _row.createCell(50).setCellValue(obj4.getInteger("loanPlatformToday"));
                    }
                    if(obj4.getInteger("loanPlatform3Days") != null)
                    {
                        _row.createCell(51).setCellValue(obj4.getInteger("loanPlatform3Days"));
                    }
                    if(obj4.getInteger("loanPlatform7Days") != null)
                    {
                        _row.createCell(52).setCellValue(obj4.getInteger("loanPlatform7Days"));
                    }
                    if(obj4.getInteger("loanPlatform15Days") != null)
                    {
                        _row.createCell(53).setCellValue(obj4.getInteger("loanPlatform15Days"));
                    }
                    if(obj4.getInteger("loanPlatform30Days") != null)
                    {
                        _row.createCell(54).setCellValue(obj4.getInteger("loanPlatform30Days"));
                    }
                    if(obj4.getInteger("loanPlatform60Days") != null)
                    {
                        _row.createCell(55).setCellValue(obj4.getInteger("loanPlatform60Days"));
                    }
                    if(obj4.getInteger("loanPlatform90Days") != null)
                    {
                        _row.createCell(56).setCellValue(obj4.getInteger("loanPlatform90Days"));
                    }
                    JSONArray arr = obj4.getJSONArray("loanPlatforms");
                    JSONObject pt;
                    StringBuilder sb = new StringBuilder();
                    if(arr != null && arr.size() > 0)
                    {
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
                    _row.createCell(57).setCellValue(sb.toString());
                }

                //优良客
                if(hr.getRepaymentCode() == 1)
                {
                    Cell cel_row = _row.createCell(58);
                    cel_row.setCellValue(0);
                    cel_row.setCellStyle(style2);
                }
                else
                {
                    Cell cel_row = _row.createCell(58);
                    cel_row.setCellValue(1);
                    cel_row.setCellStyle(style2);
                    if(obj5.getDouble("repaymentAmountMax") != null)
                    {
                        _row.createCell(59).setCellValue(obj5.getDouble("repaymentAmountMax"));
                    }
                    _row.createCell(60).setCellValue(obj5.getString("repaymentEarliestDate"));
                    _row.createCell(61).setCellValue(obj5.getString("repaymentLatestDate"));
                    if(obj5.getInteger("repaymentPlatformTotal") != null)
                    {
                        _row.createCell(62).setCellValue(obj5.getInteger("repaymentPlatformTotal"));
                    }
                    if(obj5.getInteger("repaymentPlatformToday") != null)
                    {
                        _row.createCell(63).setCellValue(obj5.getInteger("repaymentPlatformToday"));
                    }
                    if(obj5.getInteger("repaymentPlatform3Days") != null)
                    {
                        _row.createCell(64).setCellValue(obj5.getInteger("repaymentPlatform3Days"));
                    }
                    if(obj5.getInteger("repaymentPlatform7Days") != null)
                    {
                        _row.createCell(65).setCellValue(obj5.getInteger("repaymentPlatform7Days"));
                    }
                    if(obj5.getInteger("repaymentPlatform15Days") != null)
                    {
                        _row.createCell(66).setCellValue(obj5.getInteger("repaymentPlatform15Days"));
                    }
                    if(obj5.getInteger("repaymentPlatform30Days") != null)
                    {
                        _row.createCell(67).setCellValue(obj5.getInteger("repaymentPlatform30Days"));
                    }
                    if(obj5.getInteger("repaymentPlatform60Days") != null)
                    {
                        _row.createCell(68).setCellValue(obj5.getInteger("repaymentPlatform60Days"));
                    }
                    if(obj5.getInteger("repaymentPlatform90Days") != null)
                    {
                        _row.createCell(69).setCellValue(obj5.getInteger("repaymentPlatform90Days"));
                    }
                    JSONArray arr = obj5.getJSONArray("repaymentPlatforms");
                    JSONObject pt;
                    StringBuilder sb = new StringBuilder();
                    if(arr != null && arr.size() > 0)
                    {
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
                    _row.createCell(70).setCellValue(sb.toString());
                }
                c++;

            }
        }
        Row _row = sheet.createRow(c);
        /*int rows = sheet.getPhysicalNumberOfRows();*/
        System.out.println("有" + (c - 2) + "数据");
        Cell cel_ = _row.createCell(2);
        cel_.setCellValue("命中数：");

        //查询状态命中数求和
        Cell cel_row = _row.createCell(3);
        cel_row.setCellFormula("SUM(D3:D" + c + ")");
        cel_row.setCellType(CellType.FORMULA);
        XSSFFormulaEvaluator evaluator = new XSSFFormulaEvaluator(wb);
        double result = evaluator.evaluate(cel_row).getNumberValue();

        //黑名单命中数求和
        Cell cel_row0 = _row.createCell(4);
        cel_row0.setCellFormula("SUM(E3:E" + c + ")");
        cel_row0.setCellType(CellType.FORMULA);

        //常欠客命中数求和
        Cell cel_row1 = _row.createCell(5);
        cel_row1.setCellFormula("SUM(F3:F" + c + ")");
        cel_row1.setCellType(CellType.FORMULA);

        //多头客命中数求和
        Cell cel_row2 = _row.createCell(19);
        cel_row2.setCellFormula("SUM(T3:T" + c + ")");
        cel_row2.setCellType(CellType.FORMULA);

        //拒贷客命中数求和
        Cell cel_row3 = _row.createCell(32);
        cel_row3.setCellFormula("SUM(AG3:AG" + c + ")");
        cel_row3.setCellType(CellType.FORMULA);

        //通过客命中数求和
        Cell cel_row4 = _row.createCell(45);
        cel_row4.setCellFormula("SUM(AT3:AT" + c + ")");
        cel_row4.setCellType(CellType.FORMULA);

        //优良客命中数求和
        Cell cel_row5 = _row.createCell(58);
        cel_row5.setCellFormula("SUM(BG3:BG" + c + ")");
        cel_row5.setCellType(CellType.FORMULA);

        CellStyle style3 = wb.createCellStyle();
        XSSFFont font = wb.createFont();
        font.setColor(IndexedColors.RED.index);
        font.setFontName("等线");
        font.setFontHeightInPoints((short) 11);
        style3.setAlignment(HorizontalAlignment.RIGHT);
        style3.setFont(font);
        cel_.setCellStyle(style3);
        cel_row.setCellStyle(style3);
        cel_row0.setCellStyle(style3);
        cel_row1.setCellStyle(style3);
        cel_row2.setCellStyle(style3);
        cel_row3.setCellStyle(style3);
        cel_row4.setCellStyle(style3);
        cel_row5.setCellStyle(style3);

        sheet.setForceFormulaRecalculation(true);

        //命中率
        Row _row1 = sheet.createRow(c + 1);
        Cell cel_1 = _row1.createCell(3);
        if((c - 2) != 0 && result != 0)
        {
            double result1 = ((result / (c - 2)) * 100);
            System.out.println(result1);
            cel_1.setCellValue(String.format("%.2f", result1) + "%");
            cel_1.setCellStyle(style3);
        }

        OutputStream out = new FileOutputStream(outputFile);
        wb.write(out);
        out.close();
        wb.close();
    }
}