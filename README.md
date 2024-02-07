# ACC (ACI Count Collector) 

## Author and Credits 
Keyshawn Tyler (keytyler@cisco.com), Dominick Garzaniti (dgarzani@cisco.com)

## A fully functional python script that gathers count information in your ACI fabric using REST APIs
The following information this project with gather will be the following:
* Current Date
* Release Version
* EPG count 
* EP count
* Tenant count
* VRF count
* BDs count
* L3Outs
* Contracts
* Application Profile
* Leafs count
* Spines count
* APIC Count

## Requirements - Items the tool will need to function 
1. A credentials file (csv file)

## (Instructions) How to use tool
1. Clone project (git clone)
2. Run python script using py -3 ACC.py or python3 ACC.py 
3. Enter path to the csv file with APIC credentials
4. User will be prompted if an existing file with APIC information already exist. If yes, user will need to provide the path to the excel file. If no, script will continue to run. 
5. Tool will gather information from APIC with various APIC API calls. 
6. User will need to give file a name. 
7. File will be created in the directory the user is working in. 
8. Tool will ask if user wants to create another file. If yes, the tool will reset this process. If no, the tool will end. 


## Find a bug? 

If you found an issue or would like to submit an improve to this project, please submit an issue using the issues tab above. If you would like to submit a PR (pull request) with a fix, reference the issue you created!
