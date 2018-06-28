package com.guozhaotong;

import java.lang.reflect.Field;
import java.sql.*;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.List;


/**
 * 数据库操作类，包括增删改查
 */
public class MySqlUtil {
    //加载驱动
    private final String DRIVER = "com.mysql.jdbc.Driver";
    //设置url等参数
    private String URL = "jdbc:mysql://localhost:3306/testBase";
    private String userName = "root";
    private String passWord = "root";

    //定义数据库的连接
    private Connection connection;
    //定义sql语句的执行对象
    private PreparedStatement pStatement;
    //定义查询返回的结果集合
    private ResultSet resultset;

    public void setURL(String URL) {
        this.URL = URL;
    }

    public void setUserName(String userName) {
        this.userName = userName;
    }

    public void setPassWord(String passWord) {
        this.passWord = passWord;
    }

    public MySqlUtil() {
    }

    public void prepare(){
        try {
            Class.forName(DRIVER);//注册驱动
            connection = DriverManager.getConnection(URL,userName,passWord);//定义连接
            System.out.println("Connect success.");
        } catch (Exception e) {
            e.printStackTrace();
        }
    }

    public <T> List<T> multipleResultSelect(String sql, Class<T> tJavabean) {
        List<T> list = new ArrayList<T>();
        int index = 1;
        try {
            pStatement = connection.prepareStatement(sql);
            resultset = pStatement.executeQuery(sql);
            //封装resultset
            ResultSetMetaData metaData = resultset.getMetaData();//取出列的信息
            int columnLength = metaData.getColumnCount();//获取列数
            while (resultset.next()) {
                T tResult = tJavabean.newInstance();//通过反射机制创建一个对象
                for (int i = 0; i < columnLength; i++) {
                    String metaDataKey = metaData.getColumnName(i + 1);
                    Object resultsetValue = resultset.getObject(metaDataKey);
                    if (resultsetValue == null) {
                        resultsetValue = "";
                    }
                    Field field = tJavabean.getDeclaredField(metaDataKey);
                    field.setAccessible(true);
                    field.set(tResult, resultsetValue);
                }
                list.add(tResult);
            }
        } catch (SQLException | IllegalArgumentException | NoSuchFieldException | IllegalAccessException | InstantiationException | SecurityException e) {
            e.printStackTrace();
        }
        return list;
    }

    public boolean insertDeleteOperation(String sql) {
        int result = -1;
        try {
            pStatement = connection.prepareStatement(sql);
            result = pStatement.executeUpdate();//执行成功将返回大于0的数
        } catch (SQLException e) {
            e.printStackTrace();
        }
        return result > 0;
    }

    public boolean insertRow(String tabel, List<Object> params) {
        StringBuilder sql = new StringBuilder("insert into ");
        sql.append(tabel);
        sql.append(" values('");
        for (int i = 0; i < params.size() - 1; i++) {
            sql.append(params.get(i));
            sql.append("','");
        }
        sql.append(params.get(params.size() - 1));
        sql.append("')");
        return insertDeleteOperation(sql.toString());
    }

    public boolean deleteRow(String talbe, String condition) {
        return insertDeleteOperation("delete from " + talbe + " where " + condition);
    }

    public boolean deleteAllRows(String talbe) {
        return insertDeleteOperation("delete from " + talbe);
    }

    public boolean dropTable(String tableName) {
        return insertDeleteOperation("drop table " + tableName);
    }

    public boolean updateRow(String table, String setting, String condition) {
//        UPDATE Person SET FirstName = 'Fred' WHERE LastName = 'Wilson'
        return insertDeleteOperation("update " + table + " set " + setting + " where " + condition);
    }

    public <T> List<T> select(String table, String selectContent, String condition, Class<T> tJavabean){
        return multipleResultSelect("select " + selectContent + " from " + table + " where " + condition, tJavabean);
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

    public static void main(String[] args){
        MySqlUtil mySqlUtil = new MySqlUtil();
        mySqlUtil.setURL("jdbc:mysql://localhost:3306/testBase");
        mySqlUtil.setUserName("root");
        mySqlUtil.setPassWord("root");
        mySqlUtil.prepare();
        mySqlUtil.insertRow("test", Arrays.asList("guo", "24"));
        mySqlUtil.closeconnection();
    }
}