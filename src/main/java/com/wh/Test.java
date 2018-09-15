package com.wh;

import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;
import java.util.Random;

public class Test {

	public static void main(String[] args) {
		System.out.println(findNum(getData(),"域1","品牌2","型号1"));
	}
	
	
	public static Integer findNum(Map sourceData,String domainName,String brandName,String modelName){
		List list = (List) sourceData.get("colList");
		for (Object object : list) {
			Map<String,Object> domainMap=(Map<String, Object>) object;
			String targetDomainName = domainMap.get("domainName").toString();
			if(targetDomainName.equals(domainName)){
				List brandList = (List) domainMap.get("brandList");
				for (Object object2 : brandList) {
					Map<String,Object> brandMap=(Map<String, Object>) object2;
					if(brandMap.get("brandName").toString().equals(brandName)){
						List modelList=(List) brandMap.get("modelList");
						for (Object object3 : modelList) {
							Map<String,Object> modelMap=(Map<String, Object>) object3;
							if(modelMap.get("modelName").toString().equals(modelName)){
								Integer  num= (Integer) modelMap.get("modelNum");
								return num;
							}
						}
					}
				}
			}
		}
		return 0;
	}
	
	public static Map getData(){

		Random random=new Random();
		List<Map<String,Object>> list=new ArrayList();
		//按域统计  
		for(int i=0;i<3;i++){ //域code
			Map domainMap=new HashMap();
			domainMap.put("domainName", "域"+(i+1));
			List brandList=new ArrayList();
			for(int j=0;j<3;j++){ //品牌列表  这里3是需要通过域查询出来的
				Map brandMap=new HashMap();
				brandMap.put("brandName", "品牌"+(j+1));
				List modelList=new ArrayList();
				for (int k = 0; k < 2; k++) {
					int value=random.nextInt(10)+0;
					Map modelMap=new HashMap();
					modelMap.put("modelName", "型号"+(k+1));
					modelMap.put("modelNum", value);
					modelList.add(modelMap);
				}
				brandMap.put("modelList", modelList);
				brandList.add(brandMap);
			}
			domainMap.put("brandList", brandList);
			list.add(domainMap);
		}
		//按行统计
		List dbList=new ArrayList();
		int x=0;
		for(int i=0;i<6;i++){ // 6是通过group by 品牌,型号
			if(i<2){
				x=2;
			}else if(i>=2 && i<4){
				x=1;
			}else{
				x=3;
			}
			Object[] obj=new Object[]{"品牌"+x,"型号"+(random.nextInt(2)+1),
					random.nextInt(10)};
			System.out.println(obj[0].toString()+"-"+obj[1]);
			dbList.add(obj);
		}
		//继续整合数据 同品牌数对应型号列表，用于后面合并行
		Map sameBrandMap=new HashMap();
		for (int i = 0; i < dbList.size(); i++) {
			Object[] obj = (Object[]) dbList.get(i);
			String brandName = obj[0].toString();
			if(sameBrandMap.containsKey(brandName)){
				List tmpList = (List) sameBrandMap.get(brandName);
				tmpList.add(obj);
			}else{
				List tmpList=new ArrayList();
				sameBrandMap.put(brandName, tmpList);
				tmpList.add(obj);
			}
		}
		System.out.println(sameBrandMap);
		System.out.println(sameBrandMap.size());
		Map map=new HashMap();
		map.put("colList", list);
		map.put("lineList", dbList);
		map.put("lineMap", sameBrandMap);
		return map;
	
	}
	
}
