function GetFedExLabel(customerName,customerPhone,customerAddress1,customerAddress2,customerCity,customerState,customerZip,rga,shipSpeed,timestamp) {
  try{
  var testing = false;
 // customerName,customerNumber,customerAddress1,customerAddress2,customerCity,customerState,customerZip,rga,shipSpeed
/*  */ if (customerName == null)
  {
    customerName="Adepto Medical Returns";
 //   customerName="Blue Pearl Specialty & Emergency Medicine for Pets";

  }
  
  if (customerPhone == null)
  {
    customerPhone="913.261.9933";

  }
  
//  if (customerAddress1 == null)
//  {
   // customerAddress1="3120 Terrace St";
     customerAddress1="3120 Terrace St";
//  }
//  customerAddress2="";
    customerAddress2="";
    
  
//  if (customerCity == null)
//  {
   // customerCity="Kansas City";
    customerCity="Kansas City";
//  } 
  
//  if (customerState == null)
//  {
  //  customerState="MO";
    customerState="MO";
//  } 
  
//  if (customerZip == null)
 // {
   // customerZip="64111";
    customerZip="64111";
  //} 
  
  if (rga == null)
  {rga="";}
  
  if (shipSpeed == null)
  {shipSpeed = "FEDEX_GROUND";}
    
  //  if (timestamp == null)
 //   {timestamp="3/10/2018 19:06:03";}
 

  
  if (shipSpeed == null)
  {shipSpeed = "FEDEX_GROUND";}
  
  
  var xmlMessageTesting ='<soapenv:Envelope xmlns:soapenv="http://schemas.xmlsoap.org/soap/envelope/" xmlns:v19="http://fedex.com/ws/ship/v19">'
+'   <soapenv:Header/>'
+'   <soapenv:Body>'
+'<v19:ProcessShipmentRequest>'
+'  <v19:WebAuthenticationDetail>'
+'    <v19:UserCredential>'
+'      <v19:Key>vqwXtoyuAKssXDxE</v19:Key>' // testing info
+'      <v19:Password>XXXXXXXXXXXXXXXXXXXXXXXX</v19:Password>' // testing info
+'    </v19:UserCredential>'
+'  </v19:WebAuthenticationDetail>'
+'  <v19:ClientDetail>'
+'    <v19:AccountNumber>510087640</v19:AccountNumber>' // testing info
+'    <v19:MeterNumber>119018029</v19:MeterNumber>' // testing info
+'  </v19:ClientDetail>'
+'  <v19:TransactionDetail>'
+'    <v19:CustomerTransactionId>Express Return Example</v19:CustomerTransactionId>'  // testing info
+'  </v19:TransactionDetail>'
+'  <v19:Version>'
+'    <v19:ServiceId>ship</v19:ServiceId>'
+'    <v19:Major>19</v19:Major>'
+'    <v19:Intermediate>0</v19:Intermediate>'
+'    <v19:Minor>0</v19:Minor>'
+'  </v19:Version>'
+'  <v19:RequestedShipment>'
+'    <v19:ShipTimestamp>2018-03-05T12:00:00-05:00</v19:ShipTimestamp>' // what needs to go here
+'    <v19:DropoffType>BUSINESS_SERVICE_CENTER</v19:DropoffType>' // what is this
+'    <v19:ServiceType>'+shipSpeed+'</v19:ServiceType>' // what are my options
+'    <v19:PackagingType>YOUR_PACKAGING</v19:PackagingType>' // what is this
+'    <v19:Shipper>'
+'      <v19:Contact>'// update dynamically
+'        <v19:PersonName>' + customerName + '</v19:PersonName>'
+'        <v19:PhoneNumber>' + customerPhone + '</v19:PhoneNumber>'
+'      </v19:Contact>'
+'      <v19:Address>'// update dynamically
+'        <v19:StreetLines>' + customerAddress1 + '</v19:StreetLines>'
+'        <v19:StreetLines>' + customerAddress2 + '</v19:StreetLines>'
+'        <v19:City>' + customerCity + '</v19:City>'
+'        <v19:StateOrProvinceCode>' + customerState + '</v19:StateOrProvinceCode>'
+'        <v19:PostalCode>' + customerZip + '</v19:PostalCode>'
+'        <v19:CountryCode>US</v19:CountryCode>'
+'      </v19:Address>'
+'    </v19:Shipper>'
+'    <v19:Recipient>'
+'      <v19:Contact>'
+'        <v19:CompanyName>Adepto Medical</v19:CompanyName>'
+'        <v19:PhoneNumber>913.261.9933</v19:PhoneNumber>'
+'      </v19:Contact>'
+'      <v19:Address>'
+'        <v19:StreetLines>3120 Terrace St</v19:StreetLines>'
+'        <v19:StreetLines></v19:StreetLines>'
+'        <v19:City>Kansas City</v19:City>'
+'        <v19:StateOrProvinceCode>MO</v19:StateOrProvinceCode>'
+'        <v19:PostalCode>64111</v19:PostalCode>'
+'        <v19:CountryCode>US</v19:CountryCode>'
+'      </v19:Address>'
+'    </v19:Recipient>'
+'    <v19:ShippingChargesPayment>'
+'      <v19:PaymentType>SENDER</v19:PaymentType>'  /// we want to pay for it
+'      <v19:Payor>'
+'        <v19:ResponsibleParty>'
+'          <v19:AccountNumber>510087640</v19:AccountNumber>'  // testing info
+'          <v19:Contact/>'
+'        </v19:ResponsibleParty>'
+'      </v19:Payor>'
+'    </v19:ShippingChargesPayment>'
+'    <v19:SpecialServicesRequested>'
+'      <v19:SpecialServiceTypes>RETURN_SHIPMENT</v19:SpecialServiceTypes>'
+'      <v19:ReturnShipmentDetail>'
+'        <v19:ReturnType>PRINT_RETURN_LABEL</v19:ReturnType>'
+'        <v19:Rma>'
+'          <v19:Reason>Optional Reason</v19:Reason>' // return reason
+'        </v19:Rma>'
+'      </v19:ReturnShipmentDetail>'
+'    </v19:SpecialServicesRequested>'
+'    <v19:LabelSpecification>'
+'      <v19:LabelFormatType>COMMON2D</v19:LabelFormatType>'
+'      <v19:ImageType>PDF</v19:ImageType>'
+'      <v19:LabelStockType>PAPER_4X6</v19:LabelStockType>'  // what optiosn do i have
+'    </v19:LabelSpecification>'
+'    <v19:PackageCount>1</v19:PackageCount>' // This element is required if you want to process a multiple-package shipment. FedEx allows up to 99 pieces in a single transaction.
+'    <v19:RequestedPackageLineItems>'
+'      <v19:SequenceNumber>1</v19:SequenceNumber>' // what is this
+'      <v19:Weight>'// will change over time
+'        <v19:Units>LB</v19:Units>' 
+'        <v19:Value>10.0</v19:Value>'
+'      </v19:Weight>'
+'      <v19:Dimensions>'// will change over time
+'        <v19:Length>5</v19:Length>'
+'        <v19:Width>5</v19:Width>'
+'        <v19:Height>5</v19:Height>'
+'        <v19:Units>IN</v19:Units>'
+'      </v19:Dimensions>'
+'      <v19:CustomerReferences>'
+'        <v19:CustomerReferenceType>RMA_ASSOCIATION</v19:CustomerReferenceType>' // can i change this to rga
+'        <v19:Value>' + rga + '</v19:Value>' // update dynamically
+'      </v19:CustomerReferences>'
+'    </v19:RequestedPackageLineItems>'
+'  </v19:RequestedShipment>'
+'</v19:ProcessShipmentRequest>'
+'   </soapenv:Body>'
+'</soapenv:Envelope>';
    
    
var xmlMessageProduction ='<soapenv:Envelope xmlns:soapenv="http://schemas.xmlsoap.org/soap/envelope/" xmlns:v19="http://fedex.com/ws/ship/v19">'
+'   <soapenv:Header/>'
+'   <soapenv:Body>'
+'<v19:ProcessShipmentRequest>'
+'  <v19:WebAuthenticationDetail>'
+'    <v19:UserCredential>'
+'      <v19:Key>BDoRG70NlsIxocCv</v19:Key>'
+'      <v19:Password>xxxxxxxxxxxxxxxxxxxxxxxxxxxxx</v19:Password>'
+'    </v19:UserCredential>'
+'  </v19:WebAuthenticationDetail>'
+'  <v19:ClientDetail>'
+'    <v19:AccountNumber>292315619</v19:AccountNumber>'
+'    <v19:MeterNumber>112372849</v19:MeterNumber>'
+'  </v19:ClientDetail>'
+'  <v19:TransactionDetail>'
+'    <v19:CustomerTransactionId>Express Return</v19:CustomerTransactionId>' // what is this
+'  </v19:TransactionDetail>'
+'  <v19:Version>'
+'    <v19:ServiceId>ship</v19:ServiceId>'
+'    <v19:Major>19</v19:Major>'
+'    <v19:Intermediate>0</v19:Intermediate>'
+'    <v19:Minor>0</v19:Minor>'
+'  </v19:Version>'
+'  <v19:RequestedShipment>'
+'    <v19:ShipTimestamp>2018-03-05T12:00:00-05:00</v19:ShipTimestamp>' // what needs to go here
+'    <v19:DropoffType>BUSINESS_SERVICE_CENTER</v19:DropoffType>' // what is this
+'    <v19:ServiceType>'+shipSpeed+'</v19:ServiceType>' // what are my options
+'    <v19:PackagingType>YOUR_PACKAGING</v19:PackagingType>' // what is this
+'    <v19:Shipper>'
+'      <v19:Contact>'// update dynamically
+'        <v19:PersonName>' + customerName + '</v19:PersonName>'
+'        <v19:PhoneNumber>' + customerPhone + '</v19:PhoneNumber>'
+'      </v19:Contact>'
+'      <v19:Address>'// update dynamically
+'        <v19:StreetLines>' + customerAddress1 + '</v19:StreetLines>'
+'        <v19:StreetLines>' + customerAddress2 + '</v19:StreetLines>'
+'        <v19:City>' + customerCity + '</v19:City>'
+'        <v19:StateOrProvinceCode>' + customerState + '</v19:StateOrProvinceCode>'
+'        <v19:PostalCode>' + customerZip + '</v19:PostalCode>'
+'        <v19:CountryCode>US</v19:CountryCode>'
+'      </v19:Address>'
+'    </v19:Shipper>'
+'    <v19:Recipient>'
+'      <v19:Contact>'
+'        <v19:CompanyName>Adepto Medical</v19:CompanyName>'
+'        <v19:PhoneNumber>913.261.9933</v19:PhoneNumber>'
+'      </v19:Contact>'
+'      <v19:Address>'
+'        <v19:StreetLines>3120 Terrace St</v19:StreetLines>'
+'        <v19:StreetLines></v19:StreetLines>'
+'        <v19:City>Kansas City</v19:City>'
+'        <v19:StateOrProvinceCode>MO</v19:StateOrProvinceCode>'
+'        <v19:PostalCode>64111</v19:PostalCode>'
+'        <v19:CountryCode>US</v19:CountryCode>'
+'      </v19:Address>'
+'    </v19:Recipient>'
+'    <v19:ShippingChargesPayment>'
+'      <v19:PaymentType>SENDER</v19:PaymentType>'  /// we want to pay for it
+'      <v19:Payor>'
+'        <v19:ResponsibleParty>'
+'          <v19:AccountNumber>292315619</v19:AccountNumber>' 
+'          <v19:Contact/>'
+'        </v19:ResponsibleParty>'
+'      </v19:Payor>'
+'    </v19:ShippingChargesPayment>'
+'    <v19:SpecialServicesRequested>'
+'      <v19:SpecialServiceTypes>RETURN_SHIPMENT</v19:SpecialServiceTypes>'
+'      <v19:ReturnShipmentDetail>'
+'        <v19:ReturnType>PRINT_RETURN_LABEL</v19:ReturnType>'
+'        <v19:Rma>'
+'          <v19:Reason>Optional Reason</v19:Reason>' // return reason
+'        </v19:Rma>'
+'      </v19:ReturnShipmentDetail>'
+'    </v19:SpecialServicesRequested>'
+'    <v19:LabelSpecification>'
+'      <v19:LabelFormatType>COMMON2D</v19:LabelFormatType>'
+'      <v19:ImageType>PDF</v19:ImageType>'
+'      <v19:LabelStockType>PAPER_4X6</v19:LabelStockType>'  // what optiosn do i have
+'    </v19:LabelSpecification>'
+'    <v19:PackageCount>1</v19:PackageCount>' // This element is required if you want to process a multiple-package shipment. FedEx allows up to 99 pieces in a single transaction.
+'    <v19:RequestedPackageLineItems>'
+'      <v19:SequenceNumber>1</v19:SequenceNumber>' // what is this
+'      <v19:Weight>'// will change over time
+'        <v19:Units>LB</v19:Units>' 
+'        <v19:Value>10.0</v19:Value>'
+'      </v19:Weight>'
+'      <v19:Dimensions>'// will change over time
+'        <v19:Length>5</v19:Length>'
+'        <v19:Width>5</v19:Width>'
+'        <v19:Height>5</v19:Height>'
+'        <v19:Units>IN</v19:Units>'
+'      </v19:Dimensions>'
+'      <v19:CustomerReferences>'
+'        <v19:CustomerReferenceType>RMA_ASSOCIATION</v19:CustomerReferenceType>' // can i change this to rga
+'        <v19:Value>' + rga + '</v19:Value>' // update dynamically
+'      </v19:CustomerReferences>'
+'    </v19:RequestedPackageLineItems>'
+'  </v19:RequestedShipment>'
+'</v19:ProcessShipmentRequest>'
+'   </soapenv:Body>'
+'</soapenv:Envelope>';
    
    
    
  }catch(e){
  }finally{

    if (testing){
      var resultBase64 = fetchSOAP(xmlMessageTesting,testing); 
    }else{
      var resultBase64 = fetchSOAP(xmlMessageProduction,testing); 
    }
  Logger.log("resultBase64:");
  Logger.log(resultBase64);
  return resultBase64;
  }
}

function fetchSOAP(xmlMessage,testing)
{
  
    
var options =
      {
"method" : "post",
"Referrer": "AdeptoTesting",
"Host": "ws.fedex.com",
"Port": "443",
"Content-Type": "text/xml",
        "payload" : xmlMessage
      };
  Logger.log("XML:");
  Logger.log(xmlMessage);
  
  emailMe("SOAP Code ["+xmlMessage+"]", "base64 code");
  Logger.log("----------------Starting soap call-------------");
  try{
    if (testing){
      var result = UrlFetchApp.fetch("https://wsbeta.fedex.com:443/web-services", options).getContentText(); // .getContentText()
    }else{
    var result = UrlFetchApp.fetch("https://ws.fedex.com:443/web-services", options).getContentText(); // .getContentText()
    }
    
  }catch(e){
    Logger.log("error: "+e);
  }finally{
  Logger.log("======================SOAP1 responce:======================");
  Logger.log(result);
  

   Logger.log("======================base 64:======================");
  

  var base64 = extractBase64(result);
  
    Logger.log("Root: ["+base64+"]");
 //   emailMe("Code ["+base64+"]", "base64 code");
    
blob2email(base64,"12345");
    
    
  Logger.log("----------------finished with soap call-------------");
    return base64;
  }
}


function blob2email(base64,rga)
{
  if (rga == null)
  { var fileName = "FedEx_Shipping_Label.pdf";
//  var fileName = "FedEx_Shipping_Label.png";
  }
  else{
    var fileName = "FedEx_Shipping_Label_RGA_"+rga+".pdf";
  //  var fileName = "FedEx_Shipping_Label_RGA_"+rga+".png";
  }
  
  //converts base64 into a blob and saves as a pdf
    var blob  = Utilities.newBlob(Utilities.base64Decode(base64),'application/PDF' ,fileName);
 //     var blob = Utilities.newBlob(bytes, 'application/PDF', fileName);
  
  //      var blob  = Utilities.newBlob(Utilities.base64Decode(base64), 'application/PDF',fileName);
 //   var blob = Utilities.newBlob(bytes, 'application/PDF', fileName);
  

  
  
  
  var dir = DriveApp.getFolderById("1drWVHX_SluCr42LTO9k7qkpvASHPEylB");
var file = dir.createFile( blob);
  
    MailApp.sendEmail("jonathan@adeptomed.com", "test fedex", "see attached for fedex label", {attachments:[blob]});
  
}


function extractBase64(xmlString)
{
  var beforeTag= xmlString.slice(xmlString.indexOf("<Image>"), xmlString.indexOf("</Image>")); // removes everything after and including the inamge tag and everything before the opening image tag
  var afterTag = beforeTag.replace('<Image>', '');
  return afterTag;
}

// This just send me an email with the message
//    REQUIRED: string body - this is the message you want to send
function emailMe(body, sub) {
  Logger.log("Now sending email from emailMe function");
    if (sub == null) {
        var subject = "testing";
    } else {
        var subject = sub;
    }
    var message = body;
    var email = "jonathan@adeptomed.com";

    MailApp.sendEmail(email, subject, message)
}
