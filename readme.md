# **KTool: The Komprehensive PC Repair and Diagnostics Tool**

### **Overview**

KTool is a powerful PowerShell script designed to automate common IT support tasks, making PC maintenance and diagnostics more efficient. The tool can be run on a local machine or remotely on a connected host, providing a wide range of functions from clearing caches and repairing Windows components to running hardware diagnostics and updating drivers. It's a versatile utility for technicians and power users alike.

---

### **Getting Started**

To use the script, download `ktool.ps1` and save it to a local directory, such as `C:\Temp\`.

**Local Execution**

Open PowerShell as an administrator. Then, navigate to the directory where you saved the script and run it with the desired command and flag.

PowerShell

```
Set-ExecutionPolicy -ExecutionPolicy Unrestricted
C:\Temp\ktool.ps1 <command> <flag>
```

**Remote Execution**

To run KTool on a remote machine, you must have `PsExec.exe` in the same directory as the script. The remote machine must be online and accessible on the network.

PowerShell

```
Set-ExecutionPolicy -ExecutionPolicy Unrestricted
C:\Temp\ktool.ps1 remote <hostname> <command> <flag>
```

---

### **Commands**

* `repair`: Performs a full system repair, including clearing caches, running Windows repairs, updating HP drivers (if applicable), and installing Windows updates.  
* `lightrepair`: Runs only the core Windows repair functions (`DISM` and `SFC`).  
* `cache`: Clears caches for browsers (Chrome and Edge), Java, and temporary files.  
* `winupdate`: Checks for and installs available Windows updates.  
* `hpdrivers`: Uses HP Image Assistant to automatically update drivers, BIOS, and firmware on HP systems.  
* `officerep`: Runs a quick repair for Microsoft Office.  
* `slackcache`: Clears the cache for the Slack desktop application.  
* `pkill <process_name>`: Kills a specified process by name (e.g., `.\ktool-n.ps1 pkill chrome`).  
* `wlan`: Generates and displays a detailed Wi-Fi report.  
* `battery`: Generates and displays a battery health report.  
* `info`: Shows key computer and logged-in user information.  
* `adinfo <username>`: Retrieves Active Directory information for a specified user.  
* `postimage`: Runs a series of hardware and software diagnostics, including tests for disk drives, Wi-Fi, battery health, and peripherals like the keyboard and trackpad.  
* `remote`: Executes a command on a remote machine. See **Remote Execution** above for syntax.  
* `progress`: Checks the status and progress of a script running on a remote machine.  
* `notes`: Generates a text file after a remote session, creating notes for a support ticket.  
* `help`: Displays the list of available commands and flags.

---

### **Flags**

Flags are optional parameters that control the script's behavior after the main command has finished.

* `auto`: Automatically reboots the computer and deletes the script after execution.  
* `reboot`: Reboots the computer but does not delete the script.  
* `delete`: Deletes the script but does not reboot the computer.  
* No flag: The script will finish and then delete itself.

---

### **Examples**

**Local Execution**

* Run a full repair and automatically reboot and delete the script: `.\ktool-n.ps1 repair auto`  
* Clear the cache and delete the script when finished: `.\ktool-n.ps1 cache delete`  
* Run hardware diagnostics and keep the script: `.\ktool-n.ps1 postimage`

**Remote Execution**

* Run a full repair on a remote host named `PC-01`: `.\ktool-n.ps1 remote PC-01 repair auto`  
* Check the progress of a remote session on `PC-02`: `.\ktool-n.ps1 progress PC-02`  
* Kill the `Chrome` browser on a remote computer: `.\ktool-n.ps1 remote PC-03 pkill chrome`
