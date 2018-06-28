package com.guozhaotong;

import java.lang.reflect.Field;
import java.sql.*;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;


/**
 * 数据库操作类，包括增删改查
 */
public class MySqlUtil {
    //加载驱动
    private final String DRIVER = "com.mysql.jdbc.Driver";
    //设置url等参数
    private String URL = "jdbc:mysql://localhost:3306/facetextraction?user=root&password=root&useUnicode=true&characterEncoding=utf-8";
    //定义数据库的连接
    private Connection connection;
    //定义sql语句的执行对象
    private PreparedStatement pStatement;
    //定义查询返回的结果集合
    private ResultSet resultset;

    public MySqlUtil(String URL) {
        try {
            Class.forName(DRIVER);//注册驱动
            connection = DriverManager.getConnection(URL);//定义连接

        } catch (Exception e) {
            e.printStackTrace();
        }
    }

    /**
     * 完成对数据库的增删改操作
     *
     * @param sql语句
     * @param 传入的占位符，List集合
     * @return SQL语句执行成功返回true, 否则返回false
     * @throws SQLException
     */
    public boolean addDeleteModify(String sql, List<Object> params) {
        int result = -1;//设置为
        try {
            pStatement = connection.prepareStatement(sql);
            int index = 1; //从第一个开始添加
            if (params != null && !params.isEmpty()) {
                for (int i = 0; i < params.size(); i++) {
                    pStatement.setObject(index++, params.get(i));//填充占位符
                }
            }
            result = pStatement.executeUpdate();//执行成功将返回大于0的数
        } catch (SQLException e) {
            e.printStackTrace();
        }
        return result > 0;
    }

    public void res2Map(ResultSetMetaData metaDate, int columnLength, Map<String, Object> map) {
        for (int i = 0; i < columnLength; i++) {
            String metaDateKey = null;//获得列名
            try {
                metaDateKey = metaDate.getColumnName(i + 1);
                Object resultsetValue = resultset.getObject(metaDateKey);//通过列名获得值
                if (resultsetValue == null) {
                    resultsetValue = "";//转成String类型
                }
                map.put(metaDateKey, resultsetValue);//添加到map集合（以上代码是为了将从数据库返回的值转换成map的key和value）
            } catch (SQLException e) {
                e.printStackTrace();
            }
        }
    }

    /**
     * 数据库查询操作，返回单条记录
     */
    public Map<String, Object> returnSimpleResult(String sql, List<Object> params) {
        Map<String, Object> map = new HashMap<String, Object>();
        int index = 1;//从1开始设置占位符
        try {
            pStatement = connection.prepareStatement(sql);
            if (params != null && !params.isEmpty()) /*判断参数是否为空*/ {
                for (int i = 0; i < params.size(); i++) /*循环填充占位符*/ {
                    pStatement.setObject(index++, params.get(i));
                }
            }
//	        System.out.println(pStatement.toString());
            resultset = pStatement.executeQuery(sql);
            /*  将查询结果封装到map集合*/
            ResultSetMetaData metaDate = resultset.getMetaData();//获取resultSet列的信息
            int columnLength = metaDate.getColumnCount();//获得列的长度
            while (resultset.next()) {
                res2Map(metaDate, columnLength, map);
            }
        } catch (SQLException e) {
            e.printStackTrace();
        }

        return map;
    }

    /**
     * 查询数据库，返回多条记录
     */
    public List<Map<String, Object>> returnMultipleResult(String sql, List<Object> params) {
        List<Map<String, Object>> list = new ArrayList<Map<String, Object>>();
        //填充占位符
        int index = 1;
        try {
            pStatement = connection.prepareStatement(sql);
            if (params != null && !params.isEmpty()) {
                for (int i = 0; i < params.size(); i++) {
                    pStatement.setObject(index++, params.get(i));
                }
            }
            //执行SQL语句
            resultset = pStatement.executeQuery();
            //封装resultset成map类型
            ResultSetMetaData metaDate = resultset.getMetaData();//获取列信息,交给metaDate
            int columnlength = metaDate.getColumnCount();
            while (resultset.next()) {
                Map<String, Object> map = new HashMap<String, Object>();
                res2Map(metaDate, columnlength, map);
                list.add(map);
            }
        } catch (SQLException e) {
            e.printStackTrace();
        }

        return list;
    }

    /**
     * 应用反射机制返回单条记录
     *
     * @param sql
     * @param params
     * @param tJavabean，这里我用的是（SmartHome_mysql.class） javabean，我理解的是一个高度封装组件，成员为私有属性，只能
     *                                                通过set/get方法赋值和取值
     * @return 泛型
     * @throws SQLException
     */
    public <T> T returnSimpleResult_Ref(String sql, List<Object> params, Class<T> tJavabean) {
        T tResult = null;
        int index = 1;
        try {
            pStatement = connection.prepareStatement(sql);
            if (params != null && !params.isEmpty()) {
                for (int i = 0; i < params.size(); i++) {
                    pStatement.setObject(index++, params.get(i));
                }
            }
            resultset = pStatement.executeQuery(sql);
            //封装resultset
            ResultSetMetaData metaData = resultset.getMetaData();//获得列的信息
            int columnLength = metaData.getColumnCount();//获得列的长度
            while (resultset.next())//循环取值
            {
                tResult = tJavabean.newInstance();//通过反射机制创建一个实例
                function(columnLength, metaData, tResult, tJavabean);
            }
        } catch (SQLException | InstantiationException | IllegalAccessException | SecurityException e) {
            e.printStackTrace();
        }

        return tResult;
    }

    public <T> T function(int columnLength, ResultSetMetaData metaData, T tResult, Class<T> tJavabean) {
        try {
            for (int i = 0; i < columnLength; i++) {
                String metaDateKey = metaData.getColumnName(i + 1);
                Object resultsetValue = resultset.getObject(metaDateKey);
                if (resultsetValue == null) {
                    resultsetValue = "";
                }
                //获取列的属性，无论是公有。保护还是私有，都可以获取
                Field field = tJavabean.getDeclaredField(metaDateKey);
                field.setAccessible(true);//打开javabean的访问private权限
                field.set(tResult, resultsetValue);//给javabean对应的字段赋值
            }
        } catch (SQLException | IllegalAccessException | NoSuchFieldException e) {
            e.printStackTrace();
        }
        return tResult;
    }

    /**
     * 通过反射机制访问数据库，并返回多条记录
     */
    public <T> List<T> returnMultipleResult_Ref(String sql, List<Object> params, Class<T> tJavabean) {
        List<T> list = new ArrayList<T>();
        int index = 1;
        try {
            pStatement = connection.prepareStatement(sql);
            if (params != null && !params.isEmpty()) {
                for (int i = 0; i < params.size(); i++) {
                    pStatement.setObject(index, params.get(i));
                }
            }
            resultset = pStatement.executeQuery(sql);
            //封装resultset
            ResultSetMetaData metaData = resultset.getMetaData();//取出列的信息
            int columnLength = metaData.getColumnCount();//获取列数
            while (resultset.next()) {
                T tResult = tJavabean.newInstance();//通过反射机制创建一个对象
                list.add(function(columnLength, metaData, tResult, tJavabean));
            }
        } catch (SQLException | IllegalArgumentException | IllegalAccessException | InstantiationException | SecurityException e) {
            e.printStackTrace();
        }
        return list;
    }

    /**
     * 注意在finally里面执行以下方法，关闭连接
     */
    public void closeconnection() {
        if (resultset != null) {
            try {
                resultset.close();
            } catch (SQLException e) {
                e.printStackTrace();
            }
        }
        if (pStatement != null) {
            try {
                pStatement.close();
            } catch (SQLException e) {
                e.printStackTrace();
            }
        }
        if (connection != null) {
            try {
                connection.close();
            } catch (SQLException e) {
                e.printStackTrace();
            }
        }
    }
}