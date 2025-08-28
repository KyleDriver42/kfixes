# **KTool: A Swiss-Army knife for IT support.**

### **Overview**

KTool is a PowerShell script designed to automate common IT support tasks and make PC maintenance and diagnostics more efficient. It can be run on a local machine or remotely on a connected host, providing a wide range of functions for technicians.

---

### **Getting Started**

To use the script, download `ktool.ps1` and save it to a local directory, such as `C:\Temp\`.

**Local Execution**

Open PowerShell as an administrator. Then, enable Powershell script permissions, navigate to the directory where you saved the script and run it with the desired command and flag. The script automatically resets Powershell script execution permissions to "Default" when it finishes running.

PowerShell

```
Set-ExecutionPolicy -ExecutionPolicy Bypass
C:\Temp\ktool.ps1 <command> <flag>
```

**Remote Execution**

To run KTool on a remote machine, the remote machine must be online and accessible on the network. Ticket notes are generated automatically for remote sessions, and will open once the session closes.

PowerShell

```
Set-ExecutionPolicy -ExecutionPolicy Bypass
C:\Temp\ktool.ps1 remote <hostname> <command> <flag>
```

---

### **Commands**

* `repair`: Performs a full system repair, including clearing caches, running Windows repairs, updating HP drivers (if applicable), and installing Windows updates.  
* `lightrepair`: Runs only the core Windows repair functions (`DISM` and `SFC`).  
* `cache`: Clears caches for browsers (Chrome and Edge), Java, and temporary files.  
* `winupdate`: Checks for and installs available Windows updates.
* `network`: Resets Winsock, TCP/IP stack, and renews IP address   
* `hpdrivers`: Uses HP Image Assistant to automatically update drivers, BIOS, and firmware on HP systems.  
* `officerep`: Runs a quick repair for Microsoft Office.
* `printq`: Clear printer queue  
* `slackcache`: Clears the cache for the Slack desktop app.  
* `pkill <process_name>`: Kills a specified process by name (e.g., `C:\Temp\ktool.ps1 pkill chrome`).
* `errorlog`: Displays recent Application and System errors  
* `wlan`: Generates and displays a detailed Wi-Fi report.  
* `battery`: Generates and displays a battery health report.  
* `info`: Shows key computer and logged-in user information.  
* `adinfo <username>`: Retrieves Active Directory information for a specified user.  
* `postimage`: Runs a series of hardware and software diagnostics, including tests for disk drives, Wi-Fi, battery health, and peripherals like the keyboard and trackpad. Peripheral tests are skipped if run on a remote machine.
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

* Run a full repair on a remote host named `PC-01`: `C:\Temp\ktool.ps1 remote PC-01 repair auto`  
* Check the progress of a remote session on `PC-02`: `C:\Temp\ktool.ps1 progress PC-02`  
* Kill the `Chrome` browser on a remote computer: `C:\Temp\ktool.ps1 remote PC-03 pkill chrome`
