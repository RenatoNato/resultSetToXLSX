package projur.jand.program;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.Paths;
import java.sql.Connection;
import java.sql.PreparedStatement;
import java.sql.ResultSet;
import java.sql.ResultSetMetaData;
import java.sql.SQLException;
import java.text.DateFormat;
import java.text.ParseException;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.Date;
import java.util.Objects;

import org.apache.commons.lang3.time.DurationFormatUtils;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;

import projur.jand.db.DB;
import projur.jand.db.exception.FaltaParametroException;

public class ProjetoProjur {

	/**
	 * PROP�SITO - GERAR ARQUIVOS EXCEL COM VOLUMETRIA GRANDES DADOS VINDO DO BANCO DE DADOS AFIM DE 
	 * DISPONIBILIZAR RELAT�RIOS DO SISTEMA JUR�DICO PROJUR
	 * DATA DE CRIA��O 27/07/2021 DATA 
	 * DATA DA �LTIMA MODIFICA��O 12/08/2021 17:40	 * 
	 * @author rcaraujo - Renato C�zar Silva de Ara�jo
	 * @version 1.4
	 * @param args
	 * @throws SQLException
	 * @throws IOException
	 * @throws ParseException
	 */

	public static void main(String[] args) throws SQLException, ParseException {

		DateFormat df = new SimpleDateFormat("HHmmss");

//=============================== IN�CIO DO PROCESSAMENTO =============================================================

		System.out.println("\nVERS�O DO PROGRAMA [ 1.4 ]");
		String inicioProcessamento = pegaHora();
		Date iniProc = df.parse(inicioProcessamento);
		System.out.println("\n--------------------------->> INICIO DO PROCESSAMENTO [ " + pegaDataHora() + " ]\n");

//=============================== FIM DO IN�CIO DO PROCESSAMENTO ======================================================

//================================ IN�CIO PAR�METROS DO JAR ===========================================================

		String dataInicial = null;
		String dataFinal = null;
		String paramProperties = null;
		String procedure = null;
		String localGravacaoDoArquivo = null;
		String nomeArquivoGerado = null;
		String nomeDaPlanilha = null;
		int cont = 0;

		for (int i = 0; i < args.length; ++i) {

			cont++;
		}

		
		if (cont == 7) {
			
			dataInicial = args[0];
			dataFinal = args[1];
			paramProperties = args[2];
			procedure = args[3];
			localGravacaoDoArquivo = args[4];
			nomeArquivoGerado = args[5];
			nomeDaPlanilha = args[6];

		} else {

			throw new FaltaParametroException(
					"FALTA PAR�METRO FAVOR VERIFICAR A QUANTIDADE E ORDEM DOS PAR�METROS DIGITADOS");

		}

		String dirArquivo = localGravacaoDoArquivo + nomeArquivoGerado.concat("_" + pegaDataHoraArquivo() + ".xlsx");

//================================ FIM DOS PAR�METROS DO JAR ==========================================================

//================================= CONEX�O COM O BANCO DE DADOS ======================================================
		// INSERE O LOCAL DO ARQUIVO PROPERTIES PARA ORIENTA��O E CONEX�O DO BANCO
		DB.setFileProperties(paramProperties);

		// Pegando uma conex�o v�lida com o banco de dados
		Connection con = DB.getInstance().getConnection();
		System.out.println(DB.getClients());

		// EXECUTA A PROCEDURE COM PAR�METROS
		PreparedStatement ps = con.prepareStatement("? @MesAnoInicio = ? , @MesAnoFim = ?");
		ps.setString(1, procedure);
		ps.setString(2, dataInicial);
		ps.setString(3, dataFinal);

		/*
		 * ATRIBUI A EXECU��O A UM RESULSET PARA OBTERMOS O RETORNO DA
		 * CONSULTA/EXECU��O, ASSIM PODEREMOS ITERAR ENTRE AS LINHAS DA CONSULTA POIS
		 * TEREMOS UM REOTRNO
		 */
		ResultSet rs = ps.executeQuery();

//================================= FIM DAS CONFIGURA��ES DE CONEX�O COM O BANCO DE DADOS =============================

//================================= CRIA��O DE LISTAS PARA ITERA��ES ==================================================

		// LISTA CRIADA PARA OS DADOS E METADADOS DO RESULTSET
		ArrayList<String> interna = new ArrayList<String>();

//================================= FIM DAS CRIA��ES DE LISTAS PARA ITERA��ES ==================================================

//================================= CRIA��O E PROCESSAMENTO DO ARQUIVO EXCEL ==========================================

		// CRIA UM METADATA PARA OBSERMOS O LABEL(NOME DO CABE�ALHO) DE FORMA AUTOM�TICA
		// NAS CONSULTAS
		ResultSetMetaData rsmd;

		// VARI�VEL USADA PARA CONTAGEM DAS LINHAS
		// int numeroColuna;

		// CRIA��O/INSTANCIA��O DA PASTA DE TRABALHO DO EXCEL,
		// SXSSFWorkbook - CLASSE QUE SERVER PARA GERAR ARQUIVOS GRAND�ES
		SXSSFWorkbook pastaDoExcel = new SXSSFWorkbook(); // N�O USAR A CLASSE WORKBOOK APENAS SXSSFWorkbook

		// CRIA A PLANILHA DENTRO DA PASTA DE TRABALHO DO EXCEL NO CASO O ARQUIVO.XLSX E
		// DA O NOME QUE EST� ENTRE ASPAS ""
		Sheet planilha = pastaDoExcel.createSheet(nomeDaPlanilha);

		// CRIA UMA LINHA, A LINHA (0) ZERO DO CABE�ALHO
		Row cabecalho = planilha.createRow(0);

		// ATRIBU�MOS O QUE PEGAMOS DE METADAS DO RESULTSET(RETORNO DOS DADOS)
		rsmd = rs.getMetaData();
		// numeroColuna = rsmd.getColumnCount();

		// PEGA O LABEL(NOME DE CADA COLUNA DA CONSULTA) NO BANCO DE DADOS E ADICIONA A
		// LISTA CRIADA
		for (int i = 1; i <= rsmd.getColumnCount(); i++) {
			interna.add(rsmd.getColumnLabel(i));
		}

		// INSERE OS CABE�ALHOS VINDOS DOS METADATAS NAS C�LULAS DAS PLANILHA DENTRO DO
		// ARQUIVO
		for (int i = 0; i < interna.size(); i++) {
			cabecalho.createCell(i).setCellValue(interna.get(i));
		}

		// INSERIMOS O RESULTSET NAS LINHAS DO ARQUIVO, NO CASO DO EXCEL, CRIAMOS AS
		// LINHAS E DEPOIS INSERIMOS AS CELULAS
		int indiceDaLinha = 0;
		while (rs.next()) {
			Row linha = planilha.createRow(++indiceDaLinha); // CRIANDO AS LINHAS NO ARQUIVO

			for (int i = 0; i < interna.size(); i++) {

				// Cell celula = linha.createCell(i); // INSERINDO AS CELULAS NO ARQUIVO

				linha.createCell(i).setCellValue(Objects.toString(rs.getObject(interna.get(i)), ""));

			}
		}

		DB.getInstance().shutdown();

		FileOutputStream out;
		try {

			Path path = Paths.get(localGravacaoDoArquivo);

			if (Files.isDirectory(path)) {

				out = new FileOutputStream(dirArquivo);

				pastaDoExcel.write(out);
				pastaDoExcel.close();
				interna.clear();
				out.close();
				
				
				File f = new File(dirArquivo);

				boolean arquivoExiste = f.exists();

				if (arquivoExiste) {

					System.out.println("\nARQUIVO [ " + nomeArquivoGerado + " ] GERADO COM SUCESSO EM [ "
							+ path.toAbsolutePath() + " ]");

				} else {
					System.err.println("ARQUIVO N�O GERADO");
				}

			} else {
				System.err.println("DIRETORIO N�O EXISTE");
			}

		} catch (IOException e) {

			e.printStackTrace();
		}

//=============================== T�RMINO DO PROCESSAMENTO =============================================================

		String finalProcessamento = pegaHora();
		Date fimProc = df.parse(finalProcessamento);

		System.out.println("\n--------------------------->> T�RMINO DO PROCESSAMENTO [ " + pegaDataHora() + " ]");

		long duracaoProcessamento = fimProc.getTime() - iniProc.getTime();
		System.out.println("--------------------------->> DURA��O TOTAL DO PROCESSO [ "
				+ DurationFormatUtils.formatDuration(duracaoProcessamento, "HH:mm:ss") + " ]");

//=============================== FIM DO T�RMINO DO PROCESSAMENTO ======================================================	

	}

	private static String pegaDataHora() {
		DateFormat dataFormatada = new SimpleDateFormat("dd/MM/yyyy HH:mm:ss");
		Date data = new Date();
		return dataFormatada.format(data);
	}

	private static String pegaDataHoraArquivo() {
		DateFormat dataFormatada = new SimpleDateFormat("ddMMyyyy");
		Date data = new Date();
		return dataFormatada.format(data);
	}

	private static String pegaHora() {
		DateFormat dataFormatada = new SimpleDateFormat("HHmmss");
		Date data = new Date();
		return dataFormatada.format(data);
	}

}
