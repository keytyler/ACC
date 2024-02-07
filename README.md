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

## (Instructions) How to use tool
1. Clone project (git clone)
2. Run python script using py -3 ACC.py or python3 ACC.py 
3. Insert APIC IP Address 
4. Insert APIC username 
5. Insert APIC password (password will be hidden within command line)
6. Insert a name for csv (must end filename with .csv, example: Filename.csv)
7. Script will gather information using REST APIs
8. CSV file created in local directory 


## Find a bug? 

If you found an issue or would like to submit an improve to this project, please submit an issue using the issues tab above. If you would like to submit a PR (pull request) with a fix, reference the issue you created!
