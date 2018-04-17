This Vault application uses the same aliases as SAPALINK. 
More info about SAPALINK is here:
m-files://show/ACF7FC8D-2A41-43FD-BACC-FDD895870357/0-120716?object=7C684B5F-6A2A-445D-B94D-9E37AF9CD3F2

There is also some additional aliases used only in SAPDocumentConnector.
Here is a complete list of them:

Alias name				Type				Required
SAPALINK::DocumentID			Text				True
SAPALINK::SAPObjectID			Text				True
SAPALINK::SAPBusinessObjectType		Text				False
SAPALINK::DocumentType			Text				False
SAPALINK::ContentRepository		Text				True
SAPALINK::SAPObjectLink			MultiSelect from SAPObjects	True
SAPALINK::Document			Document class			True
SAPALINK::ComponentProperties		Multiline text			True
SAPALINK::DocumentProtection		Text				True
SAPALINK::ArchiveLinkVersion		Text				True
SAPALINK::DocumentDescription		Text				True
SAPALINK::WorkflowState.Archived	Workflow state			True

Aliases that are not required can be replaced by a property with a same name.

Document type and SAP Business object type can be set in a class alias in the following way:
SAPALINK::DocumentType=ZHRIAPP614;SAPALINK::SAPBusinessObjectType=PREL

Configuration is stored in Named value storage under namespace SAPDocumentConnector.
Key config contains two attributes: ConnectionString and UseFileInfo:
{
  "ConnectionString": "host=10.7.0.107; systemNumber=00;username=root;client=800;clientLanguage=EN;",
  "UseFileInfo": "false"
}
ConnectionString is passed to SAP and UseFileInfo tells if the FILENAME, CREATOR and DESCR
parameters are used when calling ARCHIV_CONNECTION_INSERT function. This requires new enough SAP
that is has SAP Note 1451769 implented.

The password for SAP connection is also stored in Named value storage. It is encrypted with the Encryptor
progrman that comes with SAPDocumentConnector. The password is in MFAdminConfiguration storage under
namespace SAPDocumentConnector. Key name is Password.
