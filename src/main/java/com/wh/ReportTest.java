package com.wh;

import java.io.File;
import java.io.FileOutputStream;
import java.math.BigInteger;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;
import java.util.Random;

import org.apache.poi.xwpf.usermodel.IBodyElement;
import org.apache.poi.xwpf.usermodel.ParagraphAlignment;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.poi.xwpf.usermodel.XWPFRun;
import org.apache.poi.xwpf.usermodel.XWPFTable;
import org.apache.poi.xwpf.usermodel.XWPFTableCell;
import org.apache.poi.xwpf.usermodel.XWPFTableRow;
import org.apache.xmlbeans.XmlCursor;
import org.apache.xmlbeans.XmlObject;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTFonts;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTHMerge;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTJc;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTOnOff;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTP;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTR;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTRPr;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTTbl;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTTblPr;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTTblWidth;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTTc;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTTcPr;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTText;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTVMerge;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.STJc;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.STMerge;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.STTblWidth;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;
import org.w3c.dom.Node;

import com.deepoove.poi.NiceXWPFDocument;
import com.deepoove.poi.XWPFTemplate;
import com.deepoove.poi.data.PictureRenderData;
import com.zte.ict.firewall.dashboardreport.data.bean.PageEchartsParams;
import com.zte.ict.firewall.dashboardreport.util.PhantomJsUtil;
import com.zte.pub.common.util.help.StringUtil;

//报表test
public class ReportTest {

	private static final Logger logger = LoggerFactory.getLogger(ReportTest.class);

	public static void main(String[] args) {
		new ReportTest().shebei_demo();
	}

	// 模板路径/保存路径
	private String savePath = "f:/echarts/";// ConfigReadUtil.getUploadConfigAsProperties("savePath");
	private String templatePath = "f:/echarts/template/";// ConfigReadUtil.getUploadConfigAsProperties("templatePath");

	public void shebei_demo() {
		logger.info("==========开始统计数据并生成报告  start");
		try {
			// 静态数据用于测试
			String surfData = "["
					+ "{\"dataTime\":1524758400000,\"dpt\":\"99999\",\"pre24h\":\"999\",\"prs\":\"99999\",\"rhu\":\"0\",\"station\":\"Y1248\",\"sunlight\":\"99999\",\"tem\":\"227\",\"temMax\":\"99999\",\"temMin\":\"99999\",\"vis\":\"99999\",\"windDAvg10mi\":\"53\",\"windSAvg10mi\":\"99999\"},"
					+ "{\"dataTime\":1524844800000,\"dpt\":\"99999\",\"pre24h\":\"399\",\"prs\":\"99999\",\"rhu\":\"0\",\"station\":\"Y1248\",\"sunlight\":\"99999\",\"tem\":\"197\",\"temMax\":\"99999\",\"temMin\":\"99999\",\"vis\":\"99999\",\"windDAvg10mi\":\"66\",\"windSAvg10mi\":\"99999\"},"
					+ "{\"dataTime\":1524931200000,\"dpt\":\"99999\",\"pre24h\":\"99999\",\"prs\":\"299\",\"rhu\":\"0\",\"station\":\"Y1248\",\"sunlight\":\"99999\",\"tem\":\"207\",\"temMax\":\"99999\",\"temMin\":\"99999\",\"vis\":\"99999\",\"windDAvg10mi\":\"62\",\"windSAvg10mi\":\"99999\"},"
					+ "{\"dataTime\":1525017600000,\"dpt\":\"99999\",\"pre24h\":\"99999\",\"prs\":\"99999\",\"rhu\":\"0\",\"station\":\"Y1248\",\"sunlight\":\"99999\",\"tem\":\"191\",\"temMax\":\"99999\",\"temMin\":\"99999\",\"vis\":\"99999\",\"windDAvg10mi\":\"49\",\"windSAvg10mi\":\"99999\"},"
					+ "{\"dataTime\":1525190400000,\"dpt\":\"99999\",\"pre24h\":\"959\",\"prs\":\"99999\",\"rhu\":\"0\",\"station\":\"Y1248\",\"sunlight\":\"\",\"tem\":\"57\",\"temMax\":\"99999\",\"temMin\":\"99999\",\"vis\":\"99999\",\"windDAvg10mi\":\"68\",\"windSAvg10mi\":\"99999\"},"
					+ "{\"dataTime\":1525276800000,\"rhu\":\"0\",\"station\":\"Y1248\",\"sunlight\":\"\",\"tem\":\"188\",\"windDAvg10mi\":\"97\",\"windSAvg10mi\":\"99999\"},{\"dataTime\":1525363200000,\"rhu\":\"0\",\"station\":\"Y1248\",\"sunlight\":\"\",\"tem\":\"216\",\"windDAvg10mi\":\"51\",\"windSAvg10mi\":\"99999\"},"
					+ "{\"dataTime\":1525449600000,\"rhu\":\"0\",\"station\":\"Y1248\",\"sunlight\":\"\",\"tem\":\"183\",\"windDAvg10mi\":\"49\",\"windSAvg10mi\":\"99999\"},"
					+ "{\"dataTime\":1525536000000,\"rhu\":\"0\",\"station\":\"Y1248\",\"sunlight\":\"\",\"tem\":\"176\",\"windDAvg10mi\":\"91\",\"windSAvg10mi\":\"99999\"},{\"dataTime\":1525622400000,\"rhu\":\"0\",\"station\":\"Y1248\",\"sunlight\":\"\",\"tem\":\"178\",\"windDAvg10mi\":\"90\",\"windSAvg10mi\":\"99999\"},"
					+ "{\"dataTime\":1525708800000,\"rhu\":\"0\",\"station\":\"Y1248\",\"sunlight\":\"\",\"tem\":\"190\",\"windDAvg10mi\":\"94\",\"windSAvg10mi\":\"99999\"},{\"dataTime\":1525795200000,\"rhu\":\"0\",\"station\":\"Y1248\",\"sunlight\":\"\",\"tem\":\"181\",\"windDAvg10mi\":\"73\",\"windSAvg10mi\":\"99999\"}]";

			List<PageEchartsParams> echartsParamsList = new ArrayList<PageEchartsParams>();
			PageEchartsParams echartsParams = new PageEchartsParams();
			echartsParams.setEchartsUrl("http://localhost:8080/echarts/testEcharts.html");
			echartsParams.setEchartsInstanceName("temEcharts");
			echartsParams.setJsFunctionName("showImg");
			echartsParams.setJsonData(surfData);
			echartsParams.setWordPlaceholder("firewall_brand_img");
			echartsParamsList.add(echartsParams);

			// echartsParamsList.add(new
			// PageEchartsParams("http://localhost:8080/echarts/testEcharts.html",
			// "yfEcharts", "showImg2",
			// null, "firewall_brand_img2"));

			// 手动调用前台js方法并填充数据返回echarts报表的base64字符串
			Map<String, File> imgFileMap = PhantomJsUtil.handlerEcharts(echartsParamsList, savePath);

			// 建立map存储所要导出到word的各种数据和图像，不能使用自己项目封装的类型，例如PageData
			Map<String, Object> dataMap = new HashMap<String, Object>();
			File imgFile = imgFileMap.get("firewall_brand_img");
			dataMap.put("firewall_brand_img", new PictureRenderData(550, 300, imgFile.getAbsolutePath()));
			// 使用poi-tl填充数据至模板
			dataMap.put("equipmentNum", 22);
			XWPFTemplate template = XWPFTemplate.compile(new File(templatePath, "设备模板.docx")).render(dataMap);
			NiceXWPFDocument xwpfDocument = template.getXWPFDocument();

			// 动态增加防火墙品牌表格
			List<Map<String, Object>> brandList = this.getFirewallBrandList();
			if (brandList != null && brandList.size() > 0) {
				String keyTabl = "${firewall_brand_table}";
				XWPFTable table = null;
				boolean addFlag = false; // 添加表格标识
				for (int i = 0; i < brandList.size(); i++) {
					if (i % 6 == 0) {// 每6个品牌生成一个新表格
						int cols; // 新表格的列数
						int index = brandList.size() - (i + 1);
						if (index < 6) {
							cols = index + 1;
							addFlag = false;
						} else {
							cols = 6;
							addFlag = true;
						}
						Map specMap = this.getSpecTablIndex(xwpfDocument, keyTabl);
						XWPFParagraph preParagraph = (XWPFParagraph) specMap.get("paragraph");
						XmlCursor newCursor = (XmlCursor) specMap.get("cursor");
						table = xwpfDocument.insertNewTbl(newCursor);// 指定位置增加table
						this.initTableRowsCols(table, 2, cols);
						this.setTableWidth(table, "8000"); // 设置表格属性
						// table = xwpfDocument.createTable(2, cols);
						// 光标后继续增加段落,分隔表格
						XmlCursor newCursor2 = preParagraph.getCTP().newCursor();
						newCursor2.toNextSibling();
						XWPFParagraph newParagraph = preParagraph.getDocument().insertNewParagraph(newCursor2);
						XWPFRun createRun = newParagraph.createRun();
						if (addFlag) {
							createRun.setText(keyTabl);
						}
						// 循环每行填充数据
						List<XWPFTableRow> rows = table.getRows();
						for (int j = 0; j < rows.size(); j++) {
							XWPFTableRow row = rows.get(j);
							List<XWPFTableCell> tableCells = row.getTableCells();
							for (int k = 0; k < tableCells.size(); k++) {
								Map<String, Object> colMap = brandList.get(k);// 单个品牌数据
								XWPFTableCell cell = tableCells.get(k);
								StringBuffer sb = new StringBuffer();
								if (j == 0) {
									sb.append(colMap.get("brandName").toString());
								} else {
									// 这里有多个型号
									List modelList = (List) colMap.get("modelList");
									for (Object model : modelList) {
										sb.append(model.toString() + "\r");
									}
								}
								this.setCellText(cell,sb.toString(),null,null);
							}
						}
					}
				}
			}
			// 2.防火墙分布表格
			Map firewallDisMap = Test.getData();//this.getFirewallBrandList2();
			List<Map<String,Object>> colList = (List) firewallDisMap.get("colList");
			Object[] lineList = (Object[]) firewallDisMap.get("lineList");
			Map<String,List> sameBrandMap = (Map<String,List>) firewallDisMap.get("lineMap"); //相同品牌下的型号集合,行合并使用
			if (colList != null && colList.size() > 0) {
				String keyTabl = "${firewall_distr_table}";
				XWPFTable table = null;
				boolean addFlag = false;
				int rowsNum = 6; // 行数 到时候去数据库查询统计
				for (int i = 0; i < colList.size(); i++) {
					if (i % 9 == 0) { // 新建table
						int cols = 0;
						int index = colList.size() - (i + 1);
						if (index < 9) {
							cols = index + 1;
							addFlag = false;
						} else {
							cols = 9;
							addFlag = true;
						}
						Map specMap = this.getSpecTablIndex(xwpfDocument, keyTabl);
						XWPFParagraph preParagraph = (XWPFParagraph) specMap.get("paragraph");
						XmlCursor newCursor = (XmlCursor) specMap.get("cursor");
						table = xwpfDocument.insertNewTbl(newCursor);// 指定位置增加table
						this.initTableRowsCols(table, rowsNum + 2, cols + 3);
						// 设置表格 填充数据
						this.setTableWidth(table, "8000");
						// 光标后继续增加段落,分隔表格
						XmlCursor newCursor2 = preParagraph.getCTP().newCursor();
						newCursor2.toNextSibling();
						XWPFParagraph newParagraph = preParagraph.getDocument().insertNewParagraph(newCursor2);
						XWPFRun createRun = newParagraph.createRun();
						if (addFlag) {
							createRun.setText(keyTabl);
						}
						List<XWPFTableRow> rows = table.getRows();
						int tmp=0;
						for (int j = 0; j < rows.size(); j++) {
							XWPFTableRow row = rows.get(j);
							List<XWPFTableCell> tableCells = row.getTableCells();
							Object[] modelArr=null;
							for (int k = 0; k < tableCells.size(); k++) {
								XWPFTableCell cell = tableCells.get(k);
								//按行处理去获取数据
								if (j == 0) { // 第一行处理
									if (k == 0 || k== 1) {
										// 合并2列
										//this.mergeCellHorizontally(table, 0, 0, 1);
										continue;
									}
									String domainName="";
									if (j == 0 && k == tableCells.size() - 1) {
										domainName = "总数";
									}else {
										Map<String, Object> domainMap = colList.get(k-2);
										domainName = (String) domainMap.get("domainName");
									}
									this.setCellText(cell, domainName,"","");
								} else if(j<rows.size()-1){
									// 一个品牌下有多少个型号 则合并几行
									Object[] dbArr = (Object[]) lineList[j-1];
									String brandName = dbArr[0].toString();
									List modelList = sameBrandMap.get(brandName);
									;
									//Map<String, Object> lineMap = lineList.get(j-1);//？？
									//String brandName = (String) lineMap.get("brandName");
									//List<Map<String,Object>> modelList = (List) lineMap.get("modelList");
									
									//跨行合并，需要定位哪几行需要合并
									if(k == 0) { //只要第一列需要合并
										if(modelList.size()>1) {
											if(j == 1) {
												this.mergeCellVertically(table, 0, j, modelList.size());
												tmp=0;
											}else {
												XWPFTableRow prevRow = rows.get(j-1); //上一行同列是否为相同品牌,如果是则不需要合并
												XWPFTableCell prevCell = prevRow.getCell(0);
												String oldName = this.getCellText(prevCell);
												if(!brandName.equals(oldName)) {
													this.mergeCellVertically(table, 0, j, modelList.size());
													tmp=0;
												}
											}
										}
										this.setCellText(cell, brandName,"","");
									}else if(k ==1) {
										modelArr = (Object[]) modelList.get(tmp);
										String modelName = modelArr[1].toString();
										this.setCellText(cell, modelName,"","");
										tmp++;
									}else if(k>=2 && k < tableCells.size()-1){ //每个品牌型号对应的域下的数量
										//Map<String, Object> map = modelList.get(tmp);
										XWPFTableRow domainRow = rows.get(0);
										XWPFTableCell domainCell = domainRow.getCell(k);
										String domainName = this.getCellText(domainCell);
										//通过域名-品牌-型号获取对应value
										Integer modelNum = Test.findNum(firewallDisMap,domainName,modelArr[0].toString(),
												modelArr[1].toString());
										if(modelNum!=null && modelNum.intValue()>0) {
											this.setCellText(cell, modelNum.toString(),"F8F8FF","4472C4");
										}else {
											this.setCellText(cell, modelNum.toString(),"","");
										}
									}else { //总数
										this.setCellText(cell, "123","F8F8FF","4472C4");
										//计算同品牌 同型号下的总值
									}
								}else if(j == rows.size()-1) {
									if(k == 0 || k == 1){
										this.setCellText(cell, "总数","","");
										continue;
									}
									//计算单个域下的总值
									this.setCellText(cell, j+"","F8F8FF","4472C4");
								}
							}
						}
						
					}
				}

			}

			File rootFile = new File(savePath);
			if (!rootFile.exists()) {
				rootFile.mkdirs();
			}
			FileOutputStream fos = new FileOutputStream(new File(rootFile, System.currentTimeMillis() + ".docx"));
			template.write(fos); // 写文件
			fos.close();
			template.close();

		} catch (Exception e) {
			e.printStackTrace();
			logger.error(e.getMessage(), e);
		}
		logger.info("========开始统计数据并生成报告  end");
	}

	/**
	 * 组装临时数据 防火墙分布
	 * 
	 * @return
	 */
	private Map getFirewallBrandList2() {
		Random random = new Random();
		List<Map<String, Object>> list = new ArrayList();
		int total = 0;
		
		List<Map<String,Object>> lineList=new ArrayList<Map<String,Object>>();
		
		for (int i = 0; i < 3; i++) { // 域的数量
			Map map = new HashMap();
			map.put("domainName", (i + 1) + "域");

			int domainTotal = 0;

			List brandList = new ArrayList();
			int rows = 0;
			for (int j = 0; j < 3; j++) { // 品牌数量
				List modelList = new ArrayList();
				for (int k = 0; k < 2; k++) { // 型号
					Map brandMap = new HashMap();
					brandMap.put("brandName", "品牌" + (j + 1));
					int value = random.nextInt(10) + 1;
					//Map modelMap = new HashMap();
					brandMap.put("modelName", "型号" + (k + 1));
					brandMap.put("modelNum", value);
					modelList.add(brandMap);
					total += value;
					domainTotal += value;
					rows++;
					if(i ==0) {
						lineList.add(brandMap);
					}
				}
				//brandMap.put("modelList", modelList);
				//brandList.add(brandMap);
				
			}
			map.put("brandList", brandList);
			map.put("domainTotal", domainTotal);
			list.add(map);
		}
		Map map=new HashMap();
		map.put("colList", list);
		
		
		//按行统计数据
		/*for (int i = 0; i < 6; i++) {
			int value = random.nextInt(10) + 1;
			Map map2=new HashMap();
			map2.put("brandName", "品牌" + (i + 1));
			map2.put("modelName", "型号" + (i + 1));
			map2.put("modelNum", value);
			map2.put("lineTotal", random.nextInt(10) + 1);
			lineList.add(map2);
		}*/
		map.put("lineList", lineList);
		return map;
	}

	/**
	 * 获取指定位置的段落和标尺
	 * 
	 * @param xwpfDocument
	 * @param string
	 * @return
	 */
	private Map getSpecTablIndex(NiceXWPFDocument xwpfDocument, String keyTabl) {
		Map map = new HashMap();
		List<XWPFParagraph> paragraphs = xwpfDocument.getParagraphs();
		for (XWPFParagraph xwpfParagraph : paragraphs) {
			List<XWPFRun> runs = xwpfParagraph.getRuns();
			for (int i = 0; i < runs.size(); i++) {
				XWPFRun xwpfRun = runs.get(i);
				String text = xwpfRun.getText(0);
				if (text != null && text.indexOf(keyTabl) != -1) {
					// xwpfRun.setText(text.replace(keyTabl, ""),0);
					xwpfParagraph.removeRun(i);
					XmlCursor newCursor = xwpfParagraph.getCTP().newCursor();
					map.put("paragraph", xwpfParagraph);
					map.put("cursor", newCursor);
					return map;
				}
			}
		}
		return null;
	}

	/**
	 * 初始化table固定的行数列数
	 * 
	 * @param table
	 * @param rows
	 * @param cols
	 */
	private void initTableRowsCols(XWPFTable table, int rows, int cols) {
		for (int i = 0; i < rows; i++) {
			XWPFTableRow tabRow = (table.getRow(i) == null) ? table.createRow() : table.getRow(i);
			for (int j = 0; j < cols; j++) {
				if (tabRow.getCell(j) == null) {
					tabRow.createCell();
				}
			}
		}
	}

	// 模拟数据,获取防火墙品牌数量
	private List<Map<String, Object>> getFirewallBrandList() {
		Random random = new Random();
		List<Map<String, Object>> brandList = new ArrayList<Map<String, Object>>();
		for (int i = 0; i < 10; i++) {
			Map<String, Object> map = new HashMap<String, Object>();
			map.put("brandName", "品牌" + (i + 1));
			List<String> modelList = new ArrayList<String>();// 型号
			// 随机增加1-5个型号
			for (int j = 0; j < random.nextInt(5) + 1; j++) {
				modelList.add("型号E12:" + random.nextInt(10) + 1);
			}
			map.put("modelList", modelList);
			brandList.add(map);
		}
		return brandList;
	}

	public void setTableWidth(XWPFTable table, String width) {
		CTTbl ttbl = table.getCTTbl();
		CTTblPr tblPr = ttbl.getTblPr() == null ? ttbl.addNewTblPr() : ttbl.getTblPr();
		CTTblWidth tblWidth = tblPr.isSetTblW() ? tblPr.getTblW() : tblPr.addNewTblW();
		CTJc cTJc = tblPr.addNewJc();
		cTJc.setVal(STJc.Enum.forString("center"));
		tblWidth.setW(new BigInteger(width));
		tblWidth.setType(STTblWidth.DXA); // 设置为固定。默认为AUTO
	}
	
	/**
	 * 设置单元格样式
	 * @param cell
	 * @param cellText
	 */
	private void setCellText(XWPFTableCell cell,String cellText,String fontColor,String bgColor){  
        CTP ctp = CTP.Factory.newInstance();
        XWPFParagraph p = new XWPFParagraph(ctp, cell);
        //p.setAlignment(ParagraphAlignment.CENTER);
        XWPFRun run = p.createRun();
        run.setText(cellText);
        if(!StringUtil.isBlank(fontColor)) {
        	run.setColor(fontColor); //字体颜色
        }
        /*CTRPr rpr = run.getCTR().isSetRPr() ? run.getCTR().getRPr() : run.getCTR().addNewRPr();  
        CTFonts fonts = rpr.isSetRFonts() ? rpr.getRFonts() : rpr.addNewRFonts();  
        fonts.setAscii("仿宋");
        fonts.setEastAsia("仿宋");
        fonts.setHAnsi("仿宋");*/
        cell.setParagraph(p);
        if(!StringUtil.isBlank(bgColor)) {
        	cell.setColor(bgColor); //背景色
        }
    }  
	
	/**
	 * 跨行合并
	 * @param table
	 * @param col
	 * @param fromRow
	 * @param toRow
	 */
	private void mergeCellVertically(XWPFTable table, int col, int fromRow, int toRow) {
        for(int rowIndex = fromRow; rowIndex <= toRow; rowIndex++){
            CTVMerge vmerge = CTVMerge.Factory.newInstance();
            if(rowIndex == fromRow){
                // The first merged cell is set with RESTART merge value
                vmerge.setVal(STMerge.RESTART);
            } else {
                // Cells which join (merge) the first one, are set with CONTINUE
                vmerge.setVal(STMerge.CONTINUE);
            }
            XWPFTableCell cell = table.getRow(rowIndex).getCell(col);
            // Try getting the TcPr. Not simply setting an new one every time.
            CTTcPr tcPr = cell.getCTTc().getTcPr();
            if (tcPr != null) {
                tcPr.setVMerge(vmerge);
            } else {
                // only set an new TcPr if there is not one already
                tcPr = CTTcPr.Factory.newInstance();
                tcPr.setVMerge(vmerge);
                cell.getCTTc().setTcPr(tcPr);
            }
        }
    }

	/**
	 * 跨列合并 目前有问题
	 * @param table
	 * @param row
	 * @param fromCol
	 * @param toCol
	 */
    private void mergeCellHorizontally(XWPFTable table, int row, int fromCol, int toCol) {
        System.out.println(1);
    	for(int colIndex = fromCol; colIndex <= toCol; colIndex++){
            CTHMerge hmerge = CTHMerge.Factory.newInstance();
            if(colIndex == fromCol){
                // The first merged cell is set with RESTART merge value
                hmerge.setVal(STMerge.RESTART);
            } else {
                // Cells which join (merge) the first one, are set with CONTINUE
                hmerge.setVal(STMerge.CONTINUE);
            }
            XWPFTableCell cell = table.getRow(row).getCell(colIndex);
            // Try getting the TcPr. Not simply setting an new one every time.
            CTTcPr tcPr = cell.getCTTc().getTcPr();
            if (tcPr != null) {
                tcPr.setHMerge(hmerge);
            } else {
                // only set an new TcPr if there is not one already
                tcPr = CTTcPr.Factory.newInstance();
                tcPr.setHMerge(hmerge);
                cell.getCTTc().setTcPr(tcPr);
            }
        }
    }
    
    /**
     * 获取指定单元格的内容
     * @param cell
     * @return
     */
    private String getCellText(XWPFTableCell cell){
    	XWPFParagraph cell_paragraph = cell.getParagraphArray(0);
		CTP ctp = cell_paragraph.getCTP();
		XmlObject  xmlObject = ctp.getRArray(0);
		XmlCursor tmpCursor = xmlObject.newCursor();
		String textValue = tmpCursor.getTextValue();
    	return textValue;
    }
    
}
