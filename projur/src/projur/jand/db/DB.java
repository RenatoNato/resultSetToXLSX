package projur.jand.db;

import java.io.FileInputStream;
import java.io.IOException;
import java.sql.Connection;
import java.sql.DriverManager;
import java.sql.SQLException;
import java.util.Properties;

/* USANDO O PADR�O SINGETON PARA CRIAR A F�BRICA DE CONEX�O, INT�NCIAR E PEGAR OS VALORES DAS 
 * PROPRIEDADES DO ARQUIVO DE CONFIGURA��ES .PROPERTIES QUE TER� A SENHA CRIPTOGRAFADA.  
 */

public class DB { // Singleton para acesso a um BD
	private static DB instance = null;
	private Connection connection = null;
	private static int clients = 0;
	private static String fileProperties;

	public static void setFileProperties(String fileProperties) {
		DB.fileProperties = fileProperties;
	}

	public static String getFileProperties() {
		return fileProperties;
	}

	private DB() { // Construtor privado, pois uso � restrito
		try {
			// Propriedades
			Properties prop = new Properties();
			prop.load(new FileInputStream(fileProperties));
			String dbDriver = prop.getProperty("db.driver");
			String dbUrl = prop.getProperty("db.url");
			String dbUser = prop.getProperty("db.user");
			String dbPwd = prop.getProperty("db.pwd");

			Class.forName(dbDriver); // passo opcional

			if (dbUser.length() != 0) { // para acesso com usu�rio e senha
				connection = DriverManager.getConnection(dbUrl, dbUser, dbPwd);
			} else { // para acesso direto (sem usu�rio e senha)
				connection = DriverManager.getConnection(dbUrl);
			}
			System.out.println("BANCO DE DADOS[ conex�o OK [ON] ]");
		} catch (ClassNotFoundException | IOException | SQLException e) {
			System.err.println(e);
		}
	}

	public static DB getInstance() { // Retorna inst�ncia �nica
		if (instance == null) {
			instance = new DB();
		}
		return instance;
	}

	public Connection getConnection() { // Retorna conex�o
		if (connection == null) {
			throw new RuntimeException("connection==null [ CONEX�O == NULA ]");
		}
		clients++;
		System.out.println("BANCO DE DADOS[ conex�o client OK [ON] ]");
		return connection;
	}

	public void shutdown() { // Efetua fechamento controlado da conex�o
		System.out.println("BANCO DE DADOS[ conex�o client desligada [OFF] ]");
		clients--;
		if (clients > 0) {
			return;
		}
		try {
			connection.close();
			instance = null;
			connection = null;
			System.out.println("BANCO DE DADOS[ conex�o fechada [OFF]]");
		} catch (SQLException sqle) {
			System.err.println(sqle);
		}
	}

	public static String getClients() {
		return "CONEX�ES ATIVAS[ " + clients + " ]";
	}

}
