package com.tcredit.imports.tables;

import java.sql.Connection;
import java.sql.DriverManager;
import java.sql.ResultSet;
import java.sql.Statement;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

/**
 * 读取mysql某数据库下表的注释信息
 * 
 * @author xxx
 */
public class MySQLTableComment {
	public static Connection getMySQLConnection() throws Exception {
		Class.forName("com.mysql.jdbc.Driver");
		Connection conn = DriverManager.getConnection("jdbc:mysql://172.19.160.74:3306/risk_platform_two", "worker",
				"Geeker4ZolZ");
		return conn;
	}

	/**
	 * 获取当前数据库下的所有表名称
	 * 
	 * @return
	 * @throws Exception
	 */
	public static List getAllTableName() throws Exception {
		List<String> tables = new ArrayList<String>();
		Connection conn = getMySQLConnection();
		Statement stmt = conn.createStatement();
		ResultSet rs = stmt.executeQuery("SHOW TABLES ");
		while (rs.next()) {
			String tableName = rs.getString(1);
			tables.add(tableName);
		}
		rs.close();
		stmt.close();
		conn.close();
		return tables;
	}

	/**
	 * 获得某表的建表语句
	 * 
	 * @param tableName
	 * @return
	 * @throws Exception
	 */
	public static Map getCommentByTableName(List tableName) throws Exception {
		Map map = new HashMap();
		Connection conn = getMySQLConnection();
		Statement stmt = conn.createStatement();
		for (int i = 0; i < tableName.size(); i++) {
			String table = (String) tableName.get(i);
			ResultSet rs = stmt.executeQuery("SHOW CREATE TABLE " + table);
			if (rs != null && rs.next()) {
				String createDDL = rs.getString(2);
				String comment = parse(createDDL);
				map.put(table, comment);
			}
			rs.close();
		}
		stmt.close();
		conn.close();
		return map;
	}

	/**
	 * 获得某表中所有字段的注释
	 * 
	 * @param tableName
	 * @return
	 * @throws Exception
	 */
	public static void getColumnCommentByTableName(List tableName) throws Exception {
		Map<String, List<String>> map = new HashMap<String, List<String>>();

		Connection conn = getMySQLConnection();
		Statement stmt = conn.createStatement();
		for (int i = 0; i < tableName.size(); i++) {
			ArrayList<String> list = new ArrayList<String>();
			String table = (String) tableName.get(i);
			ResultSet rs = stmt.executeQuery("show full columns from " + table);
			while (rs.next()) {
				if (rs.getString(2).contains("unsigned")) {
					list.add(rs.getString("Field") + "--" + rs.getString(2).replace("unsigned", "") + "--"
							+ rs.getString("Comment"));
				} else {
					list.add(rs.getString("Field") + "--" + rs.getString(2) + "--" + rs.getString("Comment"));
				}
				// System.out.println(rs.getString("Field") + "\t:\t" +
				// rs.getString(2).replace(":", "") + "\t:\t"
				// + rs.getString("Comment"));
			}
			// }
			rs.close();
			map.put("【" + table + "】", list);
		}

		stmt.close();
		conn.close();
		ImportExcel.importExcel(map);
		// return map;
	}

	/**
	 * 返回注释信息
	 * 
	 * @param all
	 * @return
	 */

	public static String parse(String all) {
		String comment = null;
		int index = all.indexOf("COMMENT='");
		if (index < 0) {
			return "";
		}
		comment = all.substring(index + 9);
		comment = comment.substring(0, comment.length() - 1);
		return comment;
	}

	public static void main(String[] args) throws Exception {
		List tables = getAllTableName();
		getColumnCommentByTableName(tables);
	}
}