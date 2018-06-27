### ExcelUtil
#### 读Excel
```
ExcelUtil poi = new ExcelUtil();
String filePath = "test.xls";
//读一个shell里面的内容
List<List<Object>> list = poi.readOneSheet(filePath, 0);
//读每个shell里面的内容
List<List<List<Object>>> lists = poi.readAllSheets(filePath);
```

#### 写Excel
```
ExcelUtil poi = new ExcelUtil();
String filePath = "test.xls";
List<List<List<String>>> res = new ArrayList<>();
List<List<String>> list = new ArrayList<>();
List<String> addedList = new ArrayList<String>();
addedList.add("我加了");
addedList.add("这一行");
addedList.add("的内容");
list.add(addedList);
//写一个工作簿，sheet名默认
poi.write(filePath, list);
//写一个工作簿，sheet名自己起
poi.write(filePath, "sheetName", list);
res.add(list);
res.add(list);
res.add(list);
//写多个工作簿，sheet名默认
poi.writeMultiSheet(filePath, res);
List<String> sheetName = Arrays.asList("sheetName0", "sheetName1", "SheetName2");
//写多个工作簿，sheet名自己起
poi.writeMultiSheet(filePath,sheetName, res);
```
