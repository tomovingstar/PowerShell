# Demonstrates how to start a service by starting the geometry service.
# This service is included with every ArcGIS Server and is stopped by default.

# To run this script, you must set the the appropriate security settings.
#  For more information about Powershell security, visit the following URL: 
#  http://www.windowsecurity.com/articles/PowerShell-Security.html


# provides serialization and deserialization functionality for JSON object
Add-Type -AssemblyName System.Web.Extensions


# Defines the entry point into the script
function main(){
  # Print some info
  echo "This is a sample script that starts the ArcGIS Server geometry service."
  
  # Ask for admin/publisher user name
  $username = read-host "Enter user name"
  $securePassword = read-host -AsSecureString "Enter password"
  
  # Ask for Admin/publisher password
  # Use Marshal Class to decode the secure string. Reference:
  #  http://msdn.microsoft.com/en-us/library/system.runtime.interopservices.marshal.aspx
  $Ptr = [System.Runtime.InteropServices.Marshal]::SecureStringToCoTaskMemUnicode($securePassword)
  $password = [System.Runtime.InteropServices.Marshal]::PtrToStringUni($Ptr)
  [System.Runtime.InteropServices.Marshal]::ZeroFreeCoTaskMemUnicode($Ptr)
  
  # Ask for servername
  $serverName = read-host "Enter server name"
  $serverPort = 6080
  
  # Get a token
  $token = token $username $password $serverName $serverPort
  if(!$token){
    write-warning "Could not generate a token with the username and password provided"
    
    return
  }
  
  # construct URL to start a service - as an example the Geometry service
  $url = "http://${serverName}:${serverPort}/arcgis/admin/services/Utilities/Geometry.GeometryServer/start"
  $parameters = "token=${token}&f=json"
  
  # Construct a HTTP request
  $http_request = New-Object -ComObject Msxml2.XMLHTTP
  $http_request.open('POST', $url, $false)
  $http_request.setRequestHeader("Content-type", "application/x-www-form-urlencoded")
  
  # Connect to URL and post parameters
  try{
    $http_request.send($parameters)
  }
  catch{
    write-warning 'There were exceptions when sending the request'
    
    return
  }
  
  # Read response
  if($http_request.status -eq 200){
    # Check that data returned is not an error object
    if(assertJsonSuccess $http_request.responseText){
      echo "Operation completed successfully!"
    }
    else{
      write-warning "Error returned by operation."
      write-warning $http_request
    }
  }
  else{
    write-warning "Error while attempting to start the service."
  }
}


# A function to generate a token given username, password and the adminURL.
function token([string]$username, [string]$password, [string]$serverName, [int]$serverPort){
  # Token URL is typically http://server[:port]/arcgis/admin/generateToken
  $url = "http://${serverName}:${serverPort}/arcgis/admin/generateToken"
  $parameters = "username=${username}&password=${password}&client=requestip&f=json"
  
  # Construct a HTTP request
  $http_request = New-Object -ComObject Msxml2.XMLHTTP
  $http_request.open('POST', $url, $false)
  $http_request.setRequestHeader("Content-type", "application/x-www-form-urlencoded")
  
  # Connect to URL and post parameters
  $token = ''
  try{
    $http_request.send($parameters)
  }
  catch{
    write-warning 'There were exceptions when sending the request'
    
    return $token
  }
  
  # Read response
  if($http_request.status -eq 200){
    # Check that data returned is not an error object
    if(assertJsonSuccess $http_request.responseText){
      # Extract the token from it
      $jsonSerializer = New-Object -TypeName System.Web.Script.Serialization.JavaScriptSerializer
      $jsonObj = $jsonSerializer.DeserializeObject($http_request.responseText)

      
      $token = $jsonObj['token']
    }
  }
  else{
    write-warning "Error encountered when retrieving an administrative token. Please check your credentials and try again."
  }
  
  $token
}


# A function that checks that the input JSON object 
#  is not an error object.  
function assertJsonSuccess([string]$data){
  $jsonSerializer = New-Object -TypeName System.Web.Script.Serialization.JavaScriptSerializer
  $jsonObj = $jsonSerializer.DeserializeObject($data)
  
  if($jsonObj['status'] -eq 'error'){
    write-warning "JSON object returns an error. Data => ${data}"
    $FALSE
  }
  else{
    $TRUE
  }
}


# Script start
main