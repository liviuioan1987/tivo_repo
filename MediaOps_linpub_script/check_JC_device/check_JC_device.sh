#!/bin/bash
rm $1_JC__applicableDeviceType.csv


for id in `cat $2` 
do

	msoName=`curl -s "http://$1.tivo.com:8085/mind/mind99?type=partnerInfoSearch&noLimit=true&partnerId=$id&levelOfDetail=high" | xmllint --format - | grep "<name>" | cut -f2 -d\> | cut -f1 -d\<`
	
	echo "MSO Name,partnerId,HeadendId,Linux" > $1_JC_applicableDeviceType_$msoName.csv    
	max=400
	echo $id
	rm $1_JC_headendId_$id.txt
      for (( offSet=0; offSet < max; offSet=offSet+50 )); 
		do
		echo $offSet
		curl -s "http://$1.tivo.com:8085/mind/mind39?type=serviceConfigurationSearch&serviceType=directTune&count=50&partnerId=$id&offset=$offSet" | xmllint --format - |  egrep -i "<headendId>" |uniq | cut -f2 -d\> | cut -f1 -d\< >> $1_JC_headendId_$id.txt
   done     
    for hid in `cat $1_JC_headendId_$id.txt` 
do
 
			applicableDeviceType=`curl -s "http://$1.tivo.com:8085/mind/mind39?type=serviceConfigurationSearch&serviceType=directTune&count=50&partnerId=$id&headendId=$hid" | xmllint --format - | grep "applicableDeviceType" | uniq | cut -f2 -d\> | cut -f1 -d\<`
			
			stbDeviceType=`curl -s "http://$1.tivo.com:8085/mind/mind39?type=serviceConfigurationSearch&serviceType=directTune&count=50&partnerId=$id&headendId=$hid" | xmllint --format - | grep "stb" | uniq | cut -f2 -d\> | cut -f1 -d\<`
			
			echo "curl -s "\"http://$1.tivo.com:8085/mind/mind39?type=serviceConfigurationSearch"&"serviceType=directTune"&"count=50"&"partnerId=$id"&"headendId=$hid\"" | xmllint --format - | grep -i "applicableDeviceType" | uniq | cut -f2 -d\> | cut -f1 -d\<"
			echo $applicableDeviceType
			
	
			 if [ -z "$applicableDeviceType" ] || [ ! -z "$stbDeviceType" ] ; then
			 isSTB="Y";
			 echo "applicableDeviceType is $applicableDeviceType";
			 else isSTB="N";
			  fi      
				#   echo "$msoName,$id,$headendId,Y" >> $1_JC_headendId_applicableDeviceType.csv
				echo "$msoName,$id,$hid,$isSTB" >> $1_JC_applicableDeviceType_$msoName.csv
			 
    done			 
done
