package excel2csv;

import com.profesorfalken.jpowershell.PowerShell;
import com.profesorfalken.jpowershell.PowerShellNotAvailableException;

public class ExecutePS {

	public static void main(String[] args) {
		try {
			PowerShell powerShell;
			powerShell = PowerShell.openSession();
			powerShell.executeCommand("powershell -ExecutionPolicy ByPass -File C:\\Users\\502362723\\Desktop\\xlsbtoxlsx.ps1 'C:\\Users\\502362723\\Desktop\\10-Mar-2017-CPS-Report-FW-10_LV' Set-ExecutionPolicy RemoteSigned");
			powerShell.close();
		} catch (PowerShellNotAvailableException e) {
			e.printStackTrace();
		}

	}

}
