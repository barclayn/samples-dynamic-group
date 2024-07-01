# samples-dynamic-group
Code samples to help you manage dynamic group membership

## PauseAll.ps1
- **Setup**: Ensure you have multiple groups with dynamic membership in unpaused state in your Microsoft 365 environment.
- **Run**: `.\PauseAll.ps1`
- **Verification**: 
- Confirm connection to Microsoft Graph with appropriate permissions.
- Verify all dynamic groups are paused in the portal.
- Check the script output for a summary of the operation.

## PauseSpecific.ps1
- **Setup**: Identify specific groups by their IDs that you want to pause.
- **Run**: `.\PauseSpecific.ps1`
- Follow the prompt to enter the IDs of the groups you want to pause.
- **Verification:**
- Confirm connection to Microsoft Graph with appropriate permissions.
- Verify the specified groups are paused in the portal.
- Check the script output for a summary of the operation.

## PauseAllExcept.ps1
- **Setup**: Identify groups by their IDs that you want to exclude from pausing.
- **Run**: `.\PauseAllExcept.ps1`
- Follow the prompt to enter the IDs of the groups you want to exclude.
- **Verification**:
- Confirm connection to Microsoft Graph with appropriate permissions.
- Verify all groups except the specified ones are paused in the portal.
- Check the script output for a summary of the operation.

## UnPauseSpecificCritical.ps1
- **Setup**: Identify specific groups by their IDs that you want to unpause.
- **Run**: `.\UnPauseSpecificCritical.ps1`
- Follow the prompt to enter the IDs of the groups you want to unpause.
- **Verification**:
- Confirm connection to Microsoft Graph with appropriate permissions.
- Verify the specified groups are unpaused in the portal.
- Check the script output for a summary of the operation.

## UnPauseNonCritical.ps1
- **Setup**: Ensure you have multiple groups with dynamic membership in paused state in your Microsoft 365 environment.
- **Run**: `.\UnPauseNonCritical.ps1`
- **Verification**:
- Confirm connection to Microsoft Graph with appropriate permissions.
- Verify some groups are unpaused in the portal.
- Check the script output for a summary of the operation.

## Error Handling and Troubleshooting
- **Connectivity Issues**: Ensure you have an active internet connection and proper credentials to connect to Microsoft Graph.
- **Permission Issues**: Verify that the account used to run the scripts has Group.ReadWrite.All permissions.
- **Script Errors**: Check the logs and error messages output by the script to identify and resolve issues. Refer to the comments and documentation within the scripts for guidance.



