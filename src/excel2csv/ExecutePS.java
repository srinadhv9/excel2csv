package excel2csv;

import com.profesorfalken.jpowershell.PowerShell;
import com.profesorfalken.jpowershell.PowerShellNotAvailableException;

public class ExecutePS {

	public static void main(String[] args) {
		try {
			PowerShell powerShell;
			powerShell = PowerShell.openSession();
			String scriptFilePath = "C:\\Users\\502362723\\Desktop\\xlsbtoxlsx.ps1";
			String sourceFile = "C:\\Users\\502362723\\Desktop\\10-Mar-2017-CPS-Report-FW-10_LV";
			String command = "powershell -ExecutionPolicy ByPass -File " + scriptFilePath + "  '" + sourceFile + "' Set-ExecutionPolicy RemoteSigned";
			powerShell.executeCommand(command);
			powerShell.close();
		} catch (PowerShellNotAvailableException e) {
			e.printStackTrace();
		}

	}

}
