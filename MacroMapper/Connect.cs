using System;
using Extensibility;
using EnvDTE;
using EnvDTE80;
using Microsoft.VisualStudio.CommandBars;
using System.Resources;
using System.Reflection;
using System.Globalization;
using Microsoft.VisualBasic;

namespace MacroMapper
{
	/// <summary>The object for implementing an Add-in.</summary>
	/// <seealso class='IDTExtensibility2' />
	public class Connect : IDTExtensibility2, IDTCommandTarget
	{
		/// <summary>Implements the constructor for the Add-in object. Place your initialization code within this method.</summary>
		public Connect()
		{
		}

		/// <summary>Implements the OnConnection method of the IDTExtensibility2 interface. Receives notification that the Add-in is being loaded.</summary>
		/// <param term='application'>Root object of the host application.</param>
		/// <param term='connectMode'>Describes how the Add-in is being loaded.</param>
		/// <param term='addInInst'>Object representing this Add-in.</param>
		/// <seealso class='IDTExtensibility2' />
		public void OnConnection(object application, ext_ConnectMode connectMode, object addInInst, ref Array custom)
		{
			_applicationObject = (DTE2)application;
			_addInInstance = (AddIn)addInInst;
			if(connectMode == ext_ConnectMode.ext_cm_UISetup)
			{
				object []contextGUIDS = new object[] { };
				Commands2 commands = (Commands2)_applicationObject.Commands;
				string toolsMenuName = "Tools";

				//Place the command on the tools menu.
				//Find the MenuBar command bar, which is the top-level command bar holding all the main menu items:
				Microsoft.VisualStudio.CommandBars.CommandBar menuBarCommandBar = ((Microsoft.VisualStudio.CommandBars.CommandBars)_applicationObject.CommandBars)["MenuBar"];

				//Find the Tools command bar on the MenuBar command bar:
				CommandBarControl toolsControl = menuBarCommandBar.Controls[toolsMenuName];
				CommandBarPopup toolsPopup = (CommandBarPopup)toolsControl;

				//This try/catch block can be duplicated if you wish to add multiple commands to be handled by your Add-in,
				//  just make sure you also update the QueryStatus/Exec method to include the new command names.
				try
				{
					//Add to the Commands collection:
                    Command commandTurnOnLineNumbers = commands.AddNamedCommand2(_addInInstance, "TurnOnLineNumbers", "TurnOnLineNumbers", "Executes the command for TurnOnLineNumbers", true, 59, ref contextGUIDS, (int)vsCommandStatus.vsCommandStatusSupported + (int)vsCommandStatus.vsCommandStatusEnabled, (int)vsCommandStyle.vsCommandStylePictAndText, vsCommandControlType.vsCommandControlTypeButton);
                    Command commandTurnOffLineNumbers = commands.AddNamedCommand2(_addInInstance, "TurnOffLineNumbers", "TurnOffLineNumbers", "Executes the command for TurnOnLineNumbers", true, 59, ref contextGUIDS, (int)vsCommandStatus.vsCommandStatusSupported + (int)vsCommandStatus.vsCommandStatusEnabled, (int)vsCommandStyle.vsCommandStylePictAndText, vsCommandControlType.vsCommandControlTypeButton);

                    //Add a control for commandTurnOnLineNumbers to the tools menu:
                    if ((commandTurnOnLineNumbers != null) && (toolsPopup != null))
                        commandTurnOnLineNumbers.AddControl(toolsPopup.CommandBar, 1);

                    //Add a control for commandTurnOffLineNumbers to the tools menu:
                    if ((commandTurnOffLineNumbers != null) && (toolsPopup != null))
                        commandTurnOffLineNumbers.AddControl(toolsPopup.CommandBar, 2);
				}
				catch(System.ArgumentException exc)
				{
					//If we are here, then the exception is probably because a command with that name
					//  already exists. If so there is no need to recreate the command and we can 
                    //  safely ignore the exception.
                    Microsoft.VisualBasic.Interaction.MsgBox(string.Format("Exception with commands: {0} {1}", exc.Message, exc.InnerException));
				}
			}
		}

		/// <summary>Implements the OnDisconnection method of the IDTExtensibility2 interface. Receives notification that the Add-in is being unloaded.</summary>
		/// <param term='disconnectMode'>Describes how the Add-in is being unloaded.</param>
		/// <param term='custom'>Array of parameters that are host application specific.</param>
		/// <seealso class='IDTExtensibility2' />
		public void OnDisconnection(ext_DisconnectMode disconnectMode, ref Array custom)
		{
		}

		/// <summary>Implements the OnAddInsUpdate method of the IDTExtensibility2 interface. Receives notification when the collection of Add-ins has changed.</summary>
		/// <param term='custom'>Array of parameters that are host application specific.</param>
		/// <seealso class='IDTExtensibility2' />		
		public void OnAddInsUpdate(ref Array custom)
		{
		}

		/// <summary>Implements the OnStartupComplete method of the IDTExtensibility2 interface. Receives notification that the host application has completed loading.</summary>
		/// <param term='custom'>Array of parameters that are host application specific.</param>
		/// <seealso class='IDTExtensibility2' />
		public void OnStartupComplete(ref Array custom)
		{
		}

		/// <summary>Implements the OnBeginShutdown method of the IDTExtensibility2 interface. Receives notification that the host application is being unloaded.</summary>
		/// <param term='custom'>Array of parameters that are host application specific.</param>
		/// <seealso class='IDTExtensibility2' />
		public void OnBeginShutdown(ref Array custom)
		{
		}
		
		/// <summary>Implements the QueryStatus method of the IDTCommandTarget interface. This is called when the command's availability is updated</summary>
		/// <param term='commandName'>The name of the command to determine state for.</param>
		/// <param term='neededText'>Text that is needed for the command.</param>
		/// <param term='status'>The state of the command in the user interface.</param>
		/// <param term='commandText'>Text requested by the neededText parameter.</param>
		/// <seealso class='Exec' />
		public void QueryStatus(string commandName, vsCommandStatusTextWanted neededText, ref vsCommandStatus status, ref object commandText)
		{
			if(neededText == vsCommandStatusTextWanted.vsCommandStatusTextWantedNone)
			{
                switch (commandName)
                {
                    case "MacroMapper.Connect.TurnOnLineNumbers":
                        status = (vsCommandStatus)vsCommandStatus.vsCommandStatusSupported | vsCommandStatus.vsCommandStatusEnabled;
                        break;
                    case "MacroMapper.Connect.TurnOffLineNumbers":
                        status = (vsCommandStatus)vsCommandStatus.vsCommandStatusSupported | vsCommandStatus.vsCommandStatusEnabled;
                        break;
                    default:
                        break;
                }
			}
		}

		/// <summary>Implements the Exec method of the IDTCommandTarget interface. This is called when the command is invoked.</summary>
		/// <param term='commandName'>The name of the command to execute.</param>
		/// <param term='executeOption'>Describes how the command should be run.</param>
		/// <param term='varIn'>Parameters passed from the caller to the command handler.</param>
		/// <param term='varOut'>Parameters passed from the command handler to the caller.</param>
		/// <param term='handled'>Informs the caller if the command was handled or not.</param>
		/// <seealso class='Exec' />
		public void Exec(string commandName, vsCommandExecOption executeOption, ref object varIn, ref object varOut, ref bool handled)
		{
			handled = false;
			if(executeOption == vsCommandExecOption.vsCommandExecOptionDoDefault)
			{
                switch (commandName)
                {
                    case "MacroMapper.Connect.TurnOnLineNumbers":
                        //Macro Sub TurnOnLineNumbers() from VisualStudio 2010 (Converted to DTE2)
                        _applicationObject.get_Properties("TextEditor", "Basic").Item("ShowLineNumbers").Value = true;
                        _applicationObject.get_Properties("TextEditor", "PlainText").Item("ShowLineNumbers").Value = true;
                        _applicationObject.get_Properties("TextEditor", "CSharp").Item("ShowLineNumbers").Value = true;
                        _applicationObject.get_Properties("TextEditor", "HTML").Item("ShowLineNumbers").Value = true;
                        _applicationObject.get_Properties("TextEditor", "C/C++").Item("ShowLineNumbers").Value = true;
                        _applicationObject.get_Properties("TextEditor", "XML").Item("ShowLineNumbers").Value = true;

                        handled = true;
                        break;

                    case "MacroMapper.Connect.TurnOffLineNumbers":
                        //Macro Sub TurnOffLineNumbers() from VisualStudio 2010 (Converted to DTE2)
                        _applicationObject.get_Properties("TextEditor", "Basic").Item("ShowLineNumbers").Value = false;
                        _applicationObject.get_Properties("TextEditor", "PlainText").Item("ShowLineNumbers").Value = false;
                        _applicationObject.get_Properties("TextEditor", "CSharp").Item("ShowLineNumbers").Value = false;
                        _applicationObject.get_Properties("TextEditor", "HTML").Item("ShowLineNumbers").Value = false;
                        _applicationObject.get_Properties("TextEditor", "C/C++").Item("ShowLineNumbers").Value = false;
                        _applicationObject.get_Properties("TextEditor", "XML").Item("ShowLineNumbers").Value = false;

                        handled = true;
                        break;

                    default:
                        break;
                }
			}
		}
		private DTE2 _applicationObject;
		private AddIn _addInInstance;
	}
}