# Office TypeLib Remediation
In some circumstances, devices that have had a previously installed 32-bit Office version may have leftover registry keys that interfere with COM and .NET functionality with a 64-bit Office version. Generally, these issues occur when the COM or .NET client is running as a 32-bit process. This project addresses remediation for these COM and .NET issues between 32-bit and 64-bit versions of Office.

## In this repository

- [Remediate-OfficeTypeLib.ps1](https://github.com/OfficeDev/Office-Typelib-Remediation/blob/main/packages/Remediate-OfficeTypeLib.ps1)

    This PowerShell script scans the registry for TypeLib keys (HKLM:\Software\WOW6432Node\Classes\TypeLib\ and HKLM:\Software\Classes\TypeLib\) for orphaned values and removes them. By default this script will run in ScanOnly mode, preventing automatic remediation. To enable remedation, set the ScanOnly parameter to $false. For example: _.\Remediate-OfficeTypeLib.ps1 -ScanOnly:$false_

- [Remediate-OfficeInterfaces.ps1](https://github.com/OfficeDev/Office-Typelib-Remediation/blob/main/packages/Remediate-OfficeInterfaces.ps1)

    This PowerShell script scans the registry for Interface keys (HKLM:\Software\WOW6432Node\Classes\Interface\) for corrupted (empty) values and removes them. By default this script will run in ScanOnly mode, preventing automatic remediation. To enable remedation, set the ScanOnly parameter to $false. For example: _.\Remediate-OfficeInterfaces.ps1 -ScanOnly:$false_

- [Configuration Baseline: Remediation for Orphaned Office TypeLib Keys.cab](https://github.com/OfficeDev/Office-Typelib-Remediation/blob/main/packages/Remediation%20for%20Orphaned%20Office%20TypeLib%20Keys.cab)

    The [Configuration Baseline](https://docs.microsoft.com/en-us/mem/configmgr/compliance/deploy-use/create-configuration-baselines#configuration-baselines) is a variation of the standalone PowerShell script, adjusted to meet the requirements for Microsoft Endpoint Configuration Manager. Once imported into Configuration Manager, the baseline can be deployed in "_Monitor_" mode (non-remediation) to identify devices that have orphaned TypeLib registry keys. With remediation enabled, the baseline will remove all known orphaned keys from the affected devices.

### Importing the Configuration Baseline
1. Download the latest version of the **Remediation for Orphaned Office TypeLib Keys.cab**.
2. Open the **Configuration Manager** console.
3. From the **Assets and Compliance** workspace, expand **Compliance Settings**.
4. Right-click on **Configuration Baselines** and select **Import Configuration Data**.
5. On the **Import Configuration Data** Wizard, click **Add** and select **Remediation for Orphaned Office TypeLib Keys.cab**. 
6. Click **Yes** on the publisher notification.
7. With the baseline selected, click **Next**.
8. On the **Summary** page, confirm the Configuration Baseline and Configuration Items match the following list. Click **Next** to complete the import.
    - Configuration Baselines (1)
      - Remediation for Orphaned Office TypeLib Keys
    - Configuration Items (1)
      - Microsoft 365 Apps - Office TypeLib Remediation
9. On the Confirmation page, click **Close**.

### Automatic Baseline Remediation
The following configuration item contains a remediation script:
  - Microsoft 365 Apps - Office TypeLib Remediation

With remediation enabled, the script will remove all known orphaned Office TypeLib registry keys.

### Deploying the Configuration Baseline
1. From the **Assets and Compliance** workspace, expand **Compliance Settings** > **Configuration Baselines**.
2. Right-click on the new baseline and select **Deploy**.
3. On the deployment dialog window, select the appropriate values and click **OK**. 
    - Check the boxes for automatic remediation as appropriate for your environment. Consider deploying without remediation first to monitor impacted devices, then enable remediation after 1-2 days.
    - Select the device collection containing devices that require detection and remediation.
    - Set the deployment schedule appropriately for your environment.

### Logging
All PowerShell scripts will log output to: _%windir%\Temp\*.log_. This can be changed using the LogFile parameter. For example: _.\Remediate-OfficeInterfaces.ps1 -LogFile C:\OfficeInterfaces.log_.

### PowerShell Execution Policy
The PowerShell scripts are provided as-is and are not signed as part of this offering. Before deploying either resource, consider the following options:

- [Sign the PowerShell scripts](https://docs.microsoft.com/en-us/powershell/module/microsoft.powershell.core/about/about_signing?view=powershell-7.1) before deployment.
- For the Configuration Manager, configure the [PowerShell execution policy](https://docs.microsoft.com/en-us/mem/configmgr/core/clients/deploy/about-client-settings#powershell-execution-policy) in Client Settings to bypass script signing requirements.

## License

Code licensed under the [MIT License](https://github.com/OfficeDev/Office-Typelib-Remediation/blob/main/LICENSE).

## Contributing

This project welcomes contributions and suggestions.  Most contributions require you to agree to a
Contributor License Agreement (CLA) declaring that you have the right to, and actually do, grant us
the rights to use your contribution. For details, visit https://cla.opensource.microsoft.com.

When you submit a pull request, a CLA bot will automatically determine whether you need to provide
a CLA and decorate the PR appropriately (e.g., status check, comment). Simply follow the instructions
provided by the bot. You will only need to do this once across all repos using our CLA.

This project has adopted the [Microsoft Open Source Code of Conduct](https://opensource.microsoft.com/codeofconduct/).
For more information see the [Code of Conduct FAQ](https://opensource.microsoft.com/codeofconduct/faq/) or
contact [opencode@microsoft.com](mailto:opencode@microsoft.com) with any additional questions or comments.

## Trademarks

This project may contain trademarks or logos for projects, products, or services. Authorized use of Microsoft 
trademarks or logos is subject to and must follow 
[Microsoft's Trademark & Brand Guidelines](https://www.microsoft.com/en-us/legal/intellectualproperty/trademarks/usage/general).
Use of Microsoft trademarks or logos in modified versions of this project must not cause confusion or imply Microsoft sponsorship.
Any use of third-party trademarks or logos are subject to those third-party's policies.
