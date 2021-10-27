package com.capgemini.appianrpa.module;

import com.capgemini.appianrpa.module.excel.ModuleExcel;
import com.novayre.jidoka.client.api.*;
import com.novayre.jidoka.client.api.annotations.Robot;
import com.novayre.jidoka.client.api.multios.IClient;
import com.novayre.jidoka.client.lowcode.IRobotVariable;

import java.io.*;
import java.nio.file.Files;
import java.nio.file.Paths;
import java.util.*;
import java.util.List;

/**
 * The Class RobotExcelTemplate.
 */
@Robot
public class RobotExcelTemplate implements IRobot {

	/** The server. */
	private IJidokaServer< ? > server;
	
	/** The client. */
	private IClient client;

	/**Content extracted from files**/
	List<String> content;
	
	/**
	 * Initialize the modules
	 */
	public void start() {
		
		server = JidokaFactory.getServer();
		client = IClient.getInstance(this);
		content = new ArrayList<>();
		server.setNumberOfItems(1);
		server.debug("Robot initialized");
	}

	/**
	 * Low-code module to get data from Excel file as a table java object (List of HashMap)
	 *
	 * @param excelFile Complete path of Excel file
	 * @param sheetExcel Sheet Name
	 * @param headerIndex Position of table header on Excel file (First row starts with 0)
	 * @param contentIndex Position of table content on Excel file (First row starts with 0)
	 * @param resultVariable Instruction to storage content form Excel
	 */
	@JidokaMethod(name = "Read Excel", description = "Method to extract data from ExcelSheet")
	public void readExcel(@JidokaParameter(defaultValue = "", name = "Excel file path") String excelFile,
								 @JidokaParameter(name = "Excel Sheet") String sheetExcel,
								 @JidokaParameter(name = "Header row") int headerIndex,
								 @JidokaParameter(name = "First row with content") int contentIndex,
						  		 @JidokaParameter(name = "Variable to storage result") String resultVariable) {
		server.setCurrentItem(1, excelFile);
		List<HashMap<String, String>> table;

		//Read Excel File
		ModuleExcel moduleExcel = new ModuleExcel();
		try {
			table = moduleExcel.redExcel(excelFile, sheetExcel, headerIndex, contentIndex);

			//Gets the map of workflow variables containing those defined on the configuration page
			Map<String, IRobotVariable> variables = server.getWorkflowVariables();

			// Gets the variable with resultVariable name
			IRobotVariable rv = variables.get(resultVariable);

			// Updates the value of resultVariable with the current value of item
			rv.setValue(table.toString());

			//Show results on trace
			table.forEach(row -> server.info(row));

			//Show variable value updated on trace
			server.info("Variable value:\n" + rv.getValue());

			//Excel file processed successfully
			server.setCurrentItemResultToOK("Lecture complete");

		} catch (IOException e) {
			//Prints error trace
			server.error(e.getMessage());
		}
	}

	/**
	 * End.
	 */
	public void end() {
		server.info("Robot finished");
	}

	@Override
	public String[] cleanUp() throws Exception {
		Files.delete(Paths.get(server.getCurrentDir()));
		return new String[0];
	}

	@Override
	public String manageException(String action, Exception exception) throws Exception{
		throw exception;
	}

}
