package com.tsn.controller;

import java.io.OutputStream;
import java.text.ParseException;
import java.util.ArrayList;
import java.util.List;

import javax.servlet.http.HttpServletRequest;
import javax.servlet.http.HttpServletResponse;

import net.sf.json.JSONObject;

import org.apache.commons.lang.StringUtils;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.springframework.stereotype.Controller;
import org.springframework.ui.Model;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.ResponseBody;

import com.tsn.util.ExcelUtil;

/**
 * 将数据库表导出为Excel表格
 * @author zhaoyuanyuan
 *
 */
@Controller
@RequestMapping("/Excel")
public class PositionController{
	
	/**
	 * 将数据检出为Excel文件
	 * @param req
	 * @param rep
	 */
	@RequestMapping("/toLeadingOut")
	public void toLeadingOut(HttpServletRequest req,HttpServletResponse rep){
		// 获取表中数据
		List<Bean> list = beanService.selectAll();
        // 设置表头
        String[] title = {"编号", "采集用户", "所在省份", "所在城市","所在区县", "设备详细地址", "经度","纬度"};
        String sheetName = "XXXX表";
        // 表内容
        String[][] content = new String[list.size()][8];
        try {
        	// 将数据存入数组
            for (int i = 0; i < list.size(); i++) {
                content[i][0] = String.valueOf(i+1);
                content[i][1] = list.get(i).getUsername();
                content[i][2] = positionService.numberToProvince(list.get(i).getProvinceId());
                content[i][3] = positionService.numberToCity(list.get(i).getCityId());
                content[i][4] = positionService.numberToCounty(list.get(i).getCountyId());
                content[i][5] = list.get(i).getSchool() + list.get(i).getAddress();
                content[i][6] = list.get(i).getLng();
                content[i][7] = list.get(i).getLat();
            }
        } catch (Exception e) {
            e.printStackTrace();
        }
        HSSFWorkbook wb = ExcelUtil.getHSSFWorkbook(sheetName, title, content, null);
        try {
        	TimeString ts = new TimeString();
			String file = ts.getTimeString();
			// 设置文件名称
			String fileName = "Table"+ file +".xls";
			OutputStream out = rep.getOutputStream();
			rep.setContentType("octets/stream");
			rep.addHeader("Content-Disposition", "attachment;filename="+fileName);
			wb.write(out);
			wb.close();
			out.flush();
			out.close();
        } catch (Exception e) {
            e.printStackTrace();
        }
	}
}
