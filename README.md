# Outlook2007CodeSamples

# _**If you're looking for the Outlook 2010 MAPI samples, go here: [Outlook 2010 MAPI Samples](https://github.com/stephenegriffin/Outlook2010CodeSamples)**_

MAPI code samples for Microsoft Outlook.

# [Click here to download the source for Outlook 2007 MAPI Samples](http://outlookmapisamples.codeplex.com/SourceControl/ListDownloadableCommits.aspx#DownloadLatest)

Documentation for these samples may be found on the [MSDN](http://msdn.microsoft.com/en-us/library/cc839588.aspx).

**Quick Details**
Date Published: 10/13/2008  
Source Language: C++  
Supported Development Platform: Microsoft Visual Studio 2008  
Download Size: 174 KB (compressed) / 659 KB (expanded)  

This download includes the following code samples:

**Sample Address Book Provider**  
This code sample implements a basic address book provider that you can use as a starting point for further customization. It implements the required features of an address book provider, as well as more advanced features such as name resolution. 

The Outlook 2007 Auxiliary Reference documents a similar sample address book provider for Visual Studio 2005. For more information about that sample address book provider, see [About the Sample Address Book Provider](http://msdn.microsoft.com/en-us/library/bb821134.aspx).

**Sample Message Store Provider**  
This code sample uses the PST provider as the back end for storing data. This wrapped PST store provider is intended to be used in conjuction with the Replication API. Most of the functions in this sample pass their arguments directly to the underlying PST provider. 

The Outlook 2007 Auxiliary Reference documents a similar sample message store provider for Visual Studio 2005. For more information about that sample wrapped PST store provider, see [About the Sample Wrapped PST Store Provider](http://msdn.microsoft.com/en-us/library/bb821132.aspx). For more information about the Replication API, see [Replication API](http://msdn.microsoft.com/en-us/library/cc160702.aspx).

**Sample Transport Provider**  
This code sample implements a shared network file system transport provider. A transport provider is a DLL that acts as an intermediary between the MAPI subsystem and one or more underlying messaging systems. You can use the code in this sample as a starting point for building a more robust transport provider.

The Outlook 2007 Auxiliary Reference documents a similar sample transport provider for Visual Studio 2005. For more information about that sample transport provider, see [About the Sample Transport Provider](http://msdn.microsoft.com/en-us/library/bb820970.aspx).
