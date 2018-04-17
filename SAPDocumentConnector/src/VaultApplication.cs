using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using MFiles.SAP;
using MFiles.VAF;
using MFiles.VAF.Common;
using MFilesAPI;
using SAP.Middleware.Connector;

namespace SAPDocumentConnector
{
	/// <summary>
	/// Simple configuration.
	/// </summary>
	public class Configuration
	{
		public string ConnectionString = "host=10.7.0.107; systemNumber=00;username=root;client=800;clientLanguage=EN;";
		public string UseFileInfo = "false";
	}

	/// <summary>
	/// Vault application to connect documents with employees in SAP.
	/// </summary>
	public class VaultApplication : VaultApplicationBase
	{
		[MFPropertyDef( Required = true )]
		public MFIdentifier DocumentIDProperty = "SAPALINK::DocumentID";
		[MFPropertyDef( Required = true )]
		public MFIdentifier SAPObjectIDProperty = "SAPALINK::SAPObjectID";
		[MFPropertyDef( Required = false )]
		public MFIdentifier SAPBusinessObjectIDProperty = "SAPALINK::SAPBusinessObjectType";
		[MFPropertyDef( Required = false )]
		public MFIdentifier DocumentTypeProperty = "SAPALINK::DocumentType";
		[ MFPropertyDef( Required = true )]
		public MFIdentifier ContentRepositoryProperty = "SAPALINK::ContentRepository";
		[MFPropertyDef( Required = true )]
		public MFIdentifier SAPObjectLinkProperty = "SAPALINK::SAPObjectLink";
		[MFClass( Required = false )]
		public MFIdentifier SAPDocumentClass = "SAPALINK::Document";
		[MFPropertyDef( Required = false )]
		public MFIdentifier ComponentPropertiesProperty = "SAPALINK::ComponentProperties";
		[MFPropertyDef( Required = true )]
		public MFIdentifier DocumentProtectionProperty = "SAPALINK::DocumentProtection";
		[MFPropertyDef( Required = true )]
		public MFIdentifier ArchiveLinkVersionProperty = "SAPALINK::ArchiveLinkVersion";
		[MFPropertyDef( Required = false )]
		public MFIdentifier DocumentDescriptionProperty = "SAPALINK::DocumentDescription";

		protected override void StartApplication()
		{
			if( MotiveHelpers.RuntimePolicyHelper.LegacyV2RuntimeEnabledSuccessfully == false )
				throw new Exception( "Enabling LegacyV2Runtime failed." );
		}

		/// <summary>
		/// Simple configuration member. MFConfiguration-attribute will cause this member to be loaded from 
		/// the named value storage from the given namespace and key.
		/// Here we override the default alias of the Configuration class with a default of the config member.
		/// The default will be written in the named value storage, if the value is missing.
		/// Use Named Value Manager to change the configurations in the named value storage.
		/// </summary>
		[MFConfiguration( "SAPDocumentConnector", "config" )]
		private Configuration config = new Configuration();

		/// <summary>
		/// Enable configuration reloading from PowerShell.
		/// </summary>
		/// <param name="env"></param>
		/// <returns></returns>
		[VaultExtensionMethod( "Reload config" )]
		private string ReloadConfig( EventHandlerEnvironment env )
		{
			NamedValues namedValues = env.Vault.NamedValueStorageOperations.GetNamedValues( MFNamedValueType.MFConfigurationValue, "SAPDocumentConnector" );

			if( namedValues.Contains( "config" ) )
			{
				config = ( Configuration ) Newtonsoft.Json.JsonConvert.DeserializeObject( ( string ) namedValues[ "config" ], typeof( Configuration ) );
			}

			return Newtonsoft.Json.JsonConvert.SerializeObject( config, Newtonsoft.Json.Formatting.Indented );
		}

		//[EventHandler( MFEventHandlerType.MFEventHandlerBeforeCheckInChangesFinalize )]  // TODO: muutetaan workflow tilaksi
		[StateAction( "SAPALINK::WorkflowState.Archived" )]
		private void LinkSAPDocument( StateEnvironment env )
		{
			SAPDataSource source = new SAPDataSource();

			// Only documents are interesting
			if( env.ObjVer.Type != ( int )MFBuiltInObjectType.MFBuiltInObjectTypeDocument )
				return;

			// Only SAP Document classes are interesting
			string objectAliases = env.Vault.ClassOperations.GetObjectClassAdmin( env.ObjVerEx.Class ).SemanticAliases.Value;
			if( !objectAliases.Contains( "SAPALINK::Document" ) )
			{
				SysUtils.ReportToEventLog( "Missing alias: SAPALINK::Document\n", System.Diagnostics.EventLogEntryType.Information );
				return;
			}

			// Map class to SAP document type. If not found, no link.
			Regex documentTypeRegex = new Regex( $"{DocumentTypeProperty.Alias}=(?<documentType>[^;]+)" );
			Match documentTypeMatch = documentTypeRegex.Match( objectAliases );
			string documentType = null;
			if( documentTypeMatch.Success )
			{
				documentType = documentTypeMatch.Groups[ "documentType" ].Value;
			}
			else if ( env.ObjVerEx.Properties.GetProperty( DocumentTypeProperty ) != null )
				documentType = env.ObjVerEx.Properties.GetProperty( DocumentTypeProperty ).GetValueAsUnlocalizedText();
			if( documentType == "" || documentType == null )
				throw new Exception( "No SAP Document type property nor SAPDocumentType alias found." );

			// SAP BusinessobjectType
			Regex objectTypeRegex = new Regex( $"{SAPBusinessObjectIDProperty.Alias}=(?<objectType>[^;]+)" );
			Match objectTypeMatch = objectTypeRegex.Match( objectAliases );
			string SAPObjectType = null;
			if( objectTypeMatch.Success )
			{
				SAPObjectType = objectTypeMatch.Groups[ "objectType" ].Value;
			}
			else if( env.ObjVerEx.Properties.GetProperty( SAPBusinessObjectIDProperty ) != null )
				SAPObjectType = env.ObjVerEx.Properties.GetProperty( SAPBusinessObjectIDProperty ).GetValueAsUnlocalizedText();
			if( SAPObjectType == "" || SAPObjectType == null )
				throw new Exception( "No SAP Business object type property nor SAPBusinessObjectType alias found." );

			// Open SAP connection
			NamedValues values = env.Vault.NamedValueStorageOperations.GetNamedValues( MFNamedValueType.MFAdminConfiguration, "SAPDocumentConnector" );
			if( !values.Contains( "Password" ) )
				throw new Exception( "The encrypted password could not be found in the NamedValuestorage." );
			source.OpenConnection( config.ConnectionString + "password=" + EncryptorLib.Encryptor.Decrypt( values[ "Password" ].ToString() ), Guid.Empty );

			// Get previous version's employee IDs.
			//PropertyValue previousSAPObjectLink = ( env.ObjVerEx.PreviousVersion != null ) ?
			//		env.ObjVerEx.PreviousVersion.Properties.GetProperty( SAPObjectLinkProperty ) : null;
			//IEnumerable<string> previousSAPObjects = ( previousSAPObjectLink != null ) ? 
			//		getSAPObjectIDs( env, previousSAPObjectLink ) : new List<string>();

			// Get current version's employee IDs.
			PropertyValue currentSAPObjectLink = env.ObjVerEx.Properties.GetProperty( SAPObjectLinkProperty );
			if( currentSAPObjectLink == null )
				throw new Exception( "No SAP SAPObject property defined." );
			IEnumerable<string> currentSAPObjects = getSAPObjectIDs( env, currentSAPObjectLink );

			//// Get deleted employee IDs.
			//IEnumerable<string> deletedIDs = previousSAPObjects.Except( currentSAPObjects );

			//// Get new employee IDs.
			//IEnumerable<string> newIDs = currentSAPObjects.Except( previousSAPObjects );

			// Remove old links
			//removeLinks( env, deletedIDs );  // TODO Doesn't work yet.

			// Create new links
			createLinks( env, documentType, SAPObjectType, currentSAPObjects, source );

			// Close connection
			try
			{
				source.CloseConnection();
			}
			catch( Exception ex )
			{
				SysUtils.ReportToEventLog( "Closing connection to SAP failed.\n" + ex.Message, System.Diagnostics.EventLogEntryType.Warning );
			}
		}

		/// <summary>
		/// Get all employee IDs from documents employee property.
		/// </summary>
		/// <param name="env"></param>
		/// <param name="employeeLink"></param>
		/// <returns></returns>
		private IEnumerable<string> getSAPObjectIDs( StateEnvironment env, PropertyValue employeeLink )
		{
			foreach( Lookup employee in employeeLink.Value.GetValueAsLookups() )
			{
				yield return getSAPObjectID( env, employee );
			}
		}

		/// <summary>
		/// Removes links between document and employees in SAP.
		/// NOT USED!
		/// </summary>
		/// <param name="env"></param>
		/// <param name="deletedIDs">List of employee IDs to be removed.</param>
		private void removeLinks( EventHandlerEnvironment env, IEnumerable<string> deletedIDs, SAPDataSource source )
		{
			foreach( string employeeID in deletedIDs )
			{
				string documentID = env.ObjVerEx.Properties.GetProperty( DocumentIDProperty ).GetValueAsLocalizedText();
				string repository = env.ObjVerEx.Properties.GetProperty( ContentRepositoryProperty ).GetValueAsUnlocalizedText();

				// Remove the link with RFC function
				IRfcFunction rfcFunction = SAPHelper.CreateFunction( "ARCHIV_DELETE_META", source.getDestination() );  // Not usable outside SAP
				rfcFunction.SetValue( "ARCHIV_ID", repository );
				rfcFunction.SetValue( "SAP_OBJECT", "PREL" );
				rfcFunction.SetValue( "ARC_DOC_ID", documentID );
				rfcFunction.SetValue( "OBJECT_ID", employeeID );
				try
				{
					string result = SAPHelper.InvokeRFC( rfcFunction, source.getDestination() );
					string message = "Removed link between document: " + documentID + "\n" +
							"and employee: " + employeeID + "\n" +
							result;
					SysUtils.ReportInfoToEventLog( message );
				}
				catch( Exception ex )
				{
					SysUtils.ReportErrorToEventLog( "SAPDocumentConnector", ex.Message, ex );
					throw new RfcInvalidParameterException( ex.Message );
				}
			}
		}

		/// <summary>
		/// Creates links between document and employees in SAP.
		/// </summary>
		/// <param name="env"></param>
		/// <param name="documentType"></param>
		/// <param name="newSAPObjects"></param>
		private void createLinks( StateEnvironment env, string documentType, string SAPObjectType, IEnumerable<string> newSAPObjects, SAPDataSource source )
		{
			foreach( string employeeID in newSAPObjects )
			{
				// Generate and set SAP properties
				string documentID = GUIDToDocumentID( env );
				env.ObjVerEx.Properties.SetProperty( DocumentIDProperty, MFDataType.MFDatatypeText, documentID );

				// Set properties required by SAP
				env.ObjVerEx.Properties.SetProperty( ComponentPropertiesProperty, MFDataType.MFDatatypeMultiLineText,
						"ComponentID=data|ContentType=application/pdf|FileID=" + GetFileID( env ) +
						"|Filename=data.pdf|ADate=" + GetSAPDateFromTS( env.ObjVerEx.GetProperty( 20 ).Value.GetValueAsTimestamp().UtcToLocalTime() ) +
						"|ATime=" + GetSAPTimeFromTS( env.ObjVerEx.GetProperty( 20 ).Value.GetValueAsTimestamp().UtcToLocalTime() ) +
						"|MDate=" + GetSAPDateFromTS( env.ObjVerEx.GetProperty( 21 ).Value.GetValueAsTimestamp().UtcToLocalTime() ) +
						"|MTime=" + GetSAPTimeFromTS( env.ObjVerEx.GetProperty( 21 ).Value.GetValueAsTimestamp().UtcToLocalTime() ) +
						"|AppVer=|Charset=" );
				env.ObjVerEx.SetProperty( DocumentProtectionProperty, MFDataType.MFDatatypeText, "rud" );
				env.ObjVerEx.SetProperty( ArchiveLinkVersionProperty, MFDataType.MFDatatypeText, "0046" );

				// Save properties
				env.ObjVerEx.SaveProperties();

				// Get SAP repository
				if( env.ObjVerEx.Properties.GetProperty( ContentRepositoryProperty ) == null )
					throw new Exception( "No SAP repository defined." );
				string repository = env.ObjVerEx.Properties.GetProperty( ContentRepositoryProperty ).GetValueAsUnlocalizedText();

					// Get filename
					string fileName = env.ObjVerEx.Title;

				// Get description
				string description = "";
				PropertyValue descProp;
				if( env.ObjVerEx.Properties.TryGetProperty( DocumentDescriptionProperty, out descProp ) )
					description = descProp.GetValueAsLocalizedText();

				// Do the link with RFC function
				IRfcFunction rfcFunction = SAPHelper.CreateFunction( "ARCHIV_CONNECTION_INSERT", source.getDestination() );
				rfcFunction.SetValue( "ARCHIV_ID", repository );
				rfcFunction.SetValue( "AR_OBJECT", documentType );
				rfcFunction.SetValue( "SAP_OBJECT", SAPObjectType );
				rfcFunction.SetValue( "ARC_DOC_ID", documentID );
				rfcFunction.SetValue( "OBJECT_ID", employeeID );

				// Is the file info in use
				if( config.UseFileInfo == "true" )
				{
					rfcFunction.SetValue( "FILENAME", fileName );
					rfcFunction.SetValue( "DESCR", description );
					rfcFunction.SetValue( "CREATOR", employeeID );
				}

				// 
				try
				{
					string result = SAPHelper.InvokeRFC( rfcFunction, source.getDestination() );
					string message = "Linked document: " + documentID + ", type: " + documentType + "\n" +
							"with SAP object: " + employeeID + ", type: " + SAPObjectType + "\n" +
							"Filename: " + fileName + ", Description: " + description + "\n" +
							result;
					SysUtils.ReportInfoToEventLog( message );
				}
				catch( Exception ex )
				{
					// Show failed FRC call in Event viewer.
					StringBuilder message = new StringBuilder()
						.AppendLine( "ARCHIV_CONNECTION_INSERT" )
						.AppendFormat( " - ARCHIV_ID: '{0}'", repository ).AppendLine()
						.AppendFormat( " - AR_OBJECT: '{0}'", documentType ).AppendLine()
						.AppendFormat( " - SAP_OBJECT: '{0}'", SAPObjectType ).AppendLine()
						.AppendFormat( " - ARC_DOC_ID: '{0}'", documentID ).AppendLine()
						.AppendFormat( " - OBJECT_ID: '{0}'", employeeID ).AppendLine();
					if( config.UseFileInfo == "true" )
					{
						message
							.AppendFormat( " - FILENAME: '{0}'", fileName ).AppendLine()
							.AppendFormat( " - DESCR: '{0}'", description ).AppendLine()
							.AppendFormat( " - CREATOR: '{0}'", employeeID ).AppendLine();
					}
					message.AppendLine().AppendLine( ex.Message );
					SysUtils.ReportErrorToEventLog( "SAPDocumentConnector", message.ToString(), ex );

					throw new RfcInvalidParameterException( ex.Message );
				}
			}
		}

		/// <summary>
		/// Gets File ID from object's first file.
		/// </summary>
		/// <param name="env"></param>
		/// <returns></returns>
		private string GetFileID( StateEnvironment env )
		{
			foreach( ObjectFile file in env.Vault.ObjectFileOperations.GetFiles( env.ObjVer ) )
			{
				return file.ID.ToString();
			}
			return null;
		}

		/// <summary>
		/// Converts Timestamp class to SAP time string.
		/// </summary>
		/// <param name="timestamp"></param>
		/// <returns></returns>
		private string GetSAPTimeFromTS( Timestamp timestamp )
		{
			return timestamp.Hour.ToString() + timestamp.Minute.ToString() + timestamp.Second.ToString();
		}

		/// <summary>
		/// Converts Timestamp class to SAP date string.
		/// </summary>
		/// <param name="ts"></param>
		/// <returns></returns>
		private string GetSAPDateFromTS( Timestamp ts )
		{
			return ts.Year.ToString() + ts.Month.ToString() + ts.Day.ToString();
		}

		/// <summary>
		/// Builds SAP document ID from M-Files GUID.
		/// </summary>
		/// <param name="env"></param>
		/// <returns></returns>
		private static string GUIDToDocumentID( StateEnvironment env )
		{
			return env.ObjVerEx.Info.ObjectGUID
					.Replace( "{", "" )
					.Replace( "}", "" )
					.Replace( "-", "" );
		}

		/// <summary>
		/// Gets SAP employee ID from M-Files lookup property.
		/// </summary>
		/// <param name="env"></param>
		/// <param name="employee"></param>
		/// <returns></returns>
		private string getSAPObjectID( StateEnvironment env, Lookup employee )
		{
			// Get SAPObject object
			ObjID objID = new ObjID();
			objID.Type = employee.ObjectType;
			objID.ID = employee.Item;
			ObjectVersionAndProperties employeeObj = env.Vault.ObjectOperations.GetLatestObjectVersionAndProperties( objID, true );

			// Get SAP SAPObject ID from SAPObject object.
			return employeeObj.Properties.GetProperty( SAPObjectIDProperty ).GetValueAsLocalizedText();
		}
	}
}