#!/bin/env node

var express = require('express');
var fs      = require('fs');
var http = require('http');
;
var path = require('path');	
var xlsx = require('xlsx');

var nodemailer = require('nodemailer');


    /*  ====================================================================== 

    /*
		
     *  Set up server IP address and port # using env variables/defaults.
     */
  function setupVariables  () {
        //  Set the environment variables we need.
       ipaddress = process.env.IP   || process.env.OPENSHIFT_NODEJS_IP || '0.0.0.0';
        port      = process.env.PORT || process.env.OPENSHIFT_NODEJS_PORT || 8080;

        if (typeof ipaddress === "undefined") {
            //  Log errors on OpenShift but continue w/ 127.0.0.1 - this
            //  allows us to run/test the app locally.
            console.warn('No OPENSHIFT_NODEJS_IP var, using 127.0.0.1');
             ipaddress = "127.0.0.1";
	           //ipaddress = '0.0.0.0';
	          //port = 8080;
        };
				
				//console.log(ipaddress+'   '+port);
    };


    /*
     *  Populate the cache.
     */
 function  populateCache  () {
        if (typeof zcache === "undefined") {
            zcache = { 'index.html': '','mgmt':'' ,'initialize':''};
        }

        //  Local cache for static content.
        zcache['index.html'] = fs.readFileSync('./index.html');
				zcache['mgmt'] = fs.readFileSync('./seatManagement.html');   
				zcache['initialize'] = fs.readFileSync('./initlz.html'); 
				
		/*		zcache['prtMartef'] = fs.readFileSync('./martefBaseHtmlToPrint.html');  
				zcache['prtRashi'] = fs.readFileSync('./rashiBaseHtmlToPrint.html');  
				zcache['prtNashim'] = fs.readFileSync('./nashimBaseHtmlToPrint.html');
				zcache['prtBase'] = fs.readFileSync('./printBaseHtml.html');
 */
			  zcache['errPasswd'] = fs.readFileSync('./errPasswd.html');   
				zcache['gizbar'] = fs.readFileSync('./gizbar.html');
				zcache['okmsg'] = fs.readFileSync('./okmsg.html');
        zcache['real_index'] = fs.readFileSync('./index_real.html');
		    zcache['moshavim'] = fs.readFileSync('./sidurUlamRahi.html');

				
    };


    /*  
     *  Retrieve entry (content) from cache.
     *  @param {string} key  Key identifying content to retrieve from cache.
     */
  function  cache_get  (key) { return zcache[key]; };


    /*
     *  terminator === the termination handler
     *  Terminate server on receipt of the specified signal.
     *  @param {string} sig  Signal to terminate on.
     */
  function terminator  (sig){
        if (typeof sig === "string") {
           console.log('%s: Received %s - terminating sample app ...',
                       Date(Date.now()), sig);
           process.exit(1);
        }
        console.log('%s: Node server stopped.', Date(Date.now()) );
    };


    /*
     *  Setup termination handlers (for exit and a list of signals).
     */
  function  setupTerminationHandlers  (){
        //  Process on exit and signals.
        process.on('exit', function() { terminator(); });

        // Removed 'SIGPIPE' from the list - bugz 852598.
        ['SIGHUP', 'SIGINT', 'SIGQUIT', 'SIGILL', 'SIGTRAP', 'SIGABRT',
         'SIGBUS', 'SIGFPE', 'SIGUSR1', 'SIGSEGV', 'SIGUSR2', 'SIGTERM'
        ].forEach(function(element, index, array) {
            process.on(element, function() { terminator(element); });
        });
    };


    /*  ================================================================  */   
    /*  App server functions (main app logic here).                       */
   //--------------------------------------------------------------------------------
  function reportErr(errText){reportAnError(errText);errorNumber=errText.substr(0,3);};
	function reportAnError(tx) {console.log(tx) ; sendErrorToKehilatArielSeatsGmail(tx);};
	function reportInputProblem(code){ reportErr(code+ txtCodes[1]);  }   //' ��� ����. ���� �������'


//--------------------------------------------------------------------------------	
	function knownName(str){  
	var rNmA = new Array(); 
	var rNm,rn,fnm_parts,startingRow,fnm_firstPossibility,fnm_secondPossibilty;
	var tempFamName,tmp,idx;
	var strParts;        
	//                   for (i=0; i<familyNames.length;i++)console.log('i='+i+' familyNames[i]='+familyNames[i]+'/');
	var strOriginalLength,indices, nuberOfPops,tempFamName,firstIdx,i,j,nextIdx,confirmedIndices;
	var firstNamesArray, nameA, nameB,bothNames;

	strParts=str.split(' ');
	strOriginalLength=strParts.length;
	
	indices=[];
	nuberOfPops=-1;
	while (strParts.length){
	    nuberOfPops++;
    	tempFamName=strParts.join(' '); //console.log('tempFamName='+tempFamName);
			firstIdx=familyNames.indexOf( tempFamName );
			if (firstIdx  == -1){ // this combination not found
			     strParts.pop();  // remove the last part of the name in case it is a first name
					 continue;  // try a shorter name
					 }  // if
			// family name found. now look for all possible families with the same family name
			  // start looking for this name from the first family name
			 indices.push([firstIdx,nuberOfPops]);  //console.log('firstIdx='+firstIdx+' nuberOfPops='+nuberOfPops); 
			 nextIdx=familyNames.indexOf( tempFamName,firstIdx+1);
			 while (nextIdx  != -1 ){
			    indices.push([nextIdx,nuberOfPops]);  //console.log('nextIdx='+nextIdx+' nuberOfPops='+nuberOfPops); 
					nextIdx=familyNames.indexOf( tempFamName,nextIdx+1);	
					}  // while nextIdx
					
			 strParts.pop();  // remove the last part of the name in case it is a first name		
			 
			 }   // while strparts.length
			 
			 // now we have all indices of possible last name
			 confirmedIndices=[];
			 //  next step == prone all indices that have different first name(s)
			  for (i=0; i<indices.length; i++){
				     nextIdx=indices[i][0];
						 nuberOfPops=indices[i][1];  // length (in tokens) of firstnames
						 strParts=str.split(' ');
						 firstNamesArray=strParts.splice(strOriginalLength-nuberOfPops,nuberOfPops);
						 
												 
						 if ( ! firstNamesArray.length){ confirmedIndices.push(nextIdx); continue};
						 // try all variations of how to generate two (or one) first names
						 for (j=0; j<firstNamesArray.length;j++){
						   trials=2;
							 confirmed=false;
						   while  ((trials) && ( ! confirmed) ){
						     nameA=firstNamesArray.splice(0,j).join(' ');
								 nameB=firstNamesArray.join(' '); 
								    if ((nameB.substr(0,1) == hebrewLetters.vav) && (trials==2) )nameB=nameB.substr(1);  
								  if (nameB <nameA){ tmp=nameA; nameA=nameB; nameB=tmp};
		                        
						  		bothNames=	sortedFirstNames[nextIdx].split('*');
									cond1= ( (nameA) && (nameA != bothNames[0]) && (nameA != bothNames[1]) );  // if specified it has to be equal to the one in the DB
									cond2= ( (nameB) && (nameB != bothNames[0]) && (nameB != bothNames[1]) );
													 
									if ( (! cond1) && ( ! cond2) ) {		confirmedIndices.push(nextIdx); confirmed=true;  break;};
									trials--;
									strParts=str.split(' ');   // restore firstNamesArray
						      firstNamesArray=strParts.splice(strOriginalLength-nuberOfPops,nuberOfPops);
									} // while
									if (confirmed )break;
									strParts=str.split(' ');   // restore firstNamesArray
						      firstNamesArray=strParts.splice(strOriginalLength-nuberOfPops,nuberOfPops);
							} // for j
					} // for i
											
				  
			if ( confirmedIndices.length != 1) { rNmA[0] = -1}else rNmA[0]=confirmedIndices[0]+firstSeatRow;
			rNmA[1]=''; 
			for (i=0; i<confirmedIndices.length;i++){
			    /*tmp=sortedFirstNames[confirmedIndices[i]].split('*');
			    nameA=tmp[0];
					nameB=tmp[1];
					if ( nameA  && nameB ){bothNames=nameA+' '+hebrewLetters.vav+nameB; } else bothNames=nameA+nameB;
					*/
					
					tmp=sortedFirstNames[confirmedIndices[i]].split('*');
					if (tmp[0]  && tmp[1]){tmp= tmp[0]+' '+hebrewLetters.vav+tmp[1]}else tmp=tmp[0]+tmp[1];
			    rNmA[1]=rNmA[1]+'$'+tmp;
			} // for i		
			
			
		return rNmA;
		
		
	};

//--------------------------------------------------------------------------------


	
 function handleInput(inPairs){  
 
 seatRelevantChangeRequest=false;
 // name
 
  tmpName=inPairs[0].split('=');  name=tmpName[1];
  if (tmpName[0]== 'familyname'){
	 rowNum=knownName(name)[0];  
	   
		 if (rowNum ==-1 ){ reportInputProblem('000'); return false;} // '�� �� ����'
		 }   
	   else { reportInputProblem('020'); return false};
			
	// email
	
		roww=rowNum.toString();
    passNotOK=false;
	tmpMail=	inPairs[1].split('=');
	 if (tmpMail[0] !='reqemail') {  reportInputProblem('002'); return false;};
		emailStr=tmpMail[1];
		emailAr= emailStr.split(',');
		email=delLeadingBlnks(emailAr[0]);
		emailPass=emailAr[1];
	  if ( email){ 
		             ptr=amudot.email+roww; 
								 crrrntEmail=delLeadingBlnks(requestedSeatsWorksheet[ptr].v);
		             
							   if(email != requestedSeatsWorksheet[ptr].v){ // console.log('forgetList[rowNum]='+forgetList[rowNum]);
								 passNotOK= ( ! emailPass)  || (forgetList[rowNum].split('$')[0] != emailPass);
								     if ( (crrrntEmail) && passNotOK ) { // changing email but password not supplied or wrong
							          console.log('bad attempt to change passcode. row='+roww+ ' old email='+requestedSeatsWorksheet[ptr].v+
												          ' new attemted email='+email);
												return false;					
												 
										 }  //changing email but password not supplied or wrong
										 else {requestedSeatsWorksheet[ptr].v=email;}  // ok change email
							      
							  }    // if email !=
					}  // if email
		 	
	
 //  phone
 
 tmpPhone=inPairs[2].split('=');	
    if (tmpPhone[0]=='reqPhone') {
		phne=tmpPhone[1];
		if (phne){ ptr=amudot.phone+roww;    requestedSeatsWorksheet[ptr].v=phne;}
		}
				 else {	reportInputProblem('003'); return false;};
		
	
// address

tmpAddress=inPairs[3].split('=');	
    if (tmpAddress[0]=='reqaddress') {
		address=tmpAddress[1];
		if (address){ ptr=amudot.addr+roww;    requestedSeatsWorksheet[ptr].v=address;}
		} else {	reportInputProblem('004'); return false;};
	
// menRosh

tmpCount=inPairs[4].split('=');	
    if ((tmpCount[0]!='menRosh') || (isNaN(tmpCount[1]))) { 	reportInputProblem('005'); return false;} 
		   else {
			  ptr=amudot.menRosh+roww;
			    if(tmpCount[1]) {  numMenRosh=tmpCount[1];} else {numMenRosh=0};    
					if (requestedSeatsWorksheet[ptr].v !=numMenRosh) {
					      requestedSeatsWorksheet[ptr].v =numMenRosh;
								seatRelevantChangeRequest=true;
								}
		        };
// menKipur

tmpCount=inPairs[5].split('=');	  
 if ((tmpCount[0]!='menKipur') || (isNaN(tmpCount[1])))  	{reportInputProblem('006'); return false;}
      else {
			ptr=amudot.menKipur+roww; 
			if(tmpCount[1]) {numMenKipur=tmpCount[1];} else {numMenKipur=0};
			if (requestedSeatsWorksheet[ptr].v !=numMenKipur) {
					      requestedSeatsWorksheet[ptr].v =numMenKipur;
								seatRelevantChangeRequest=true;
								}
			
 };
 		  numMen=Math.max(numMenKipur,numMenRosh);
// womenRosh

tmpCount=inPairs[6].split('=');	  
 if ((tmpCount[0]!='womenRosh') || (isNaN(tmpCount[1])))  {reportInputProblem('007'); return false;}
    else {
		ptr=amudot.womenRosh+roww;
		 if(tmpCount[1]) {	 numWomnRosh=tmpCount[1];} else {numWomnRosh=0};
		 if (requestedSeatsWorksheet[ptr].v !=numWomnRosh) {
					      requestedSeatsWorksheet[ptr].v =numWomnRosh;
								seatRelevantChangeRequest=true;
								}
		 
		};
		
// womenKipur

tmpCount=inPairs[7].split('=');	  
 if ((tmpCount[0]!='womenKipur') || (isNaN(tmpCount[1]))) { reportInputProblem('0011'); return false;}
    else {
		ptr=amudot.womenKipur+roww;
		 if(tmpCount[1]) {	 numWomnKipur=tmpCount[1];} else {numWomnKipur=0};
		 if (requestedSeatsWorksheet[ptr].v !=numWomnKipur) {
					      requestedSeatsWorksheet[ptr].v =numWomnKipur;
								seatRelevantChangeRequest=true;
								}
		 
		
		 numWomn=Math.max(numWomnRosh,numWomnKipur); 
		};
		
// choose a Minyan for women    
  
    tmpMinyan=inPairs[8].split('=');  
		reqMinyanData=inPairs[9].split('='); 
				if ( (tmpMinyan[0] != 'reqminyanW') || (reqMinyanData[0] != 'reqMinyanDataW')|| isNaN(tmpMinyan[1] )
				              || ( tmpMinyan[1]<0 )|| ( (tmpMinyan[1]>5)&& (tmpMinyan[1] != 9 ))  )  {reportInputProblem('008'); return false;}
					
					NtmpMinyan=Number(tmpMinyan[1]); 
					setUlam( NtmpMinyan,amudot.nashimMuadaf+roww) ;
					if (tmpMinyan[1] != 9 ){ reqMinyan[0]= txtCodes[3+NtmpMinyan];	}
					     else	reqMinyan[0]='9';
          reqMinyan[1]=reqMinyanData[1];			
							
			   ptr=amudot.preferedMinyanW+roww;  
			   if (requestedSeatsWorksheet[ptr].v !=reqMinyan[0]) {
					      requestedSeatsWorksheet[ptr].v =reqMinyan[0];
								seatRelevantChangeRequest=true;
								}
					ptr=amudot.preferedExplanationW+roww;    requestedSeatsWorksheet[ptr].v= reqMinyan[1]    
	


			
// choose a Minyan for men
  
    tmpMinyan=inPairs[10].split('=');  
		reqMinyanData=inPairs[11].split('=');   
				if ( (tmpMinyan[0] != 'reqminyanM') || (reqMinyanData[0] != 'reqMinyanDataM')|| isNaN(tmpMinyan[1] )
				              || ( tmpMinyan[1]<0 )|| ( (tmpMinyan[1]>5)&& (tmpMinyan[1] != 9 ))  )  {reportInputProblem('008'); return false;}
					
					NtmpMinyan=Number(tmpMinyan[1]); 
						setUlam( NtmpMinyan,amudot.gvarimMuadaf+roww);   
					if (tmpMinyan[1] != 9 ){ reqMinyan[0]= txtCodes[9+NtmpMinyan];	}
					     else	reqMinyan[0]='9';		
          reqMinyan[1]=reqMinyanData[1];			
							
			ptr=amudot.preferedMinyanM+roww; 
			if (requestedSeatsWorksheet[ptr].v !=reqMinyan[0]) {
					      requestedSeatsWorksheet[ptr].v =reqMinyan[0];
								seatRelevantChangeRequest=true;
								}
			ptr=amudot.preferedExplanationM+roww;    requestedSeatsWorksheet[ptr].v= reqMinyan[1]    				
			
			
			
			
			
			
			
	//   more comments
							
	tmpCmnt=inPairs[12].split('=');
		if (tmpCmnt[0] != 'RmoreComments'){reportInputProblem('009'); return false};
				ptr=amudot.cmnts+roww;    requestedSeatsWorksheet[ptr].v=tmpCmnt[1];    // trnslTxt(tmpCmnt[1])
		
		
	//seats
	
	


	tmpSeats=inPairs[13].split('=');    
	if (tmpSeats[0] != 'seatsString'){reportInputProblem('010'); return false}
			
		ptr=amudot.markedSeats+roww;   
		seats=tmpSeats[1].split('+');
		seats=seats.sort(sortOrder);  if( ! delLeadingBlnks(tmpSeats[1]) ){ seats=[];};
		lastSeatsRequest=(requestedSeatsWorksheet[ptr].v).split('+');  if( ! delLeadingBlnks(requestedSeatsWorksheet[ptr].v) ) lastSeatsRequest=[];
		changeInRequest=false;
		if(seats.length != lastSeatsRequest.length){
		                changeInRequest=true; }
					else {
					  for  (i=0; i<seats.length; i++)if (seats[i].split('_')[0] != 	lastSeatsRequest[i].split('_')[0])changeInRequest=true; 
						}
			
			if(	changeInRequest ){
			  						
										seatRelevantChangeRequest=true;
										}
										
		tmppSeats=	seats.join('+');							
		 requestedSeatsWorksheet[ptr].v=tmppSeats
		
		
		
		
		countM=0;    countW=0;
		for(i=0; i<seats.length; i++) {
		  k=Number(seats[i].split('_')[0] ); 
			if( isWoman[k]){countW++;} else {countM++;};
			}
			
			ptr=amudot.numberMarkedMen+roww;  requestedSeatsWorksheet[ptr].v=countM.toString();
			ptr=amudot.numberMarkedWomen+roww;  requestedSeatsWorksheet[ptr].v=countW.toString();
			
	                               		
		
 incompatibilty='&0';	  
 if( ( ( countM)	&&	(countM<numMen) ) || ( ( countW) && (countW<numWomn) ) )incompatibilty='&1';
	
		
	//tashlum
	tmpTashlum=inPairs[14].split('=');    
	if (tmpTashlum[0] != 'tashlum'){reportInputProblem('014'); return false}
			
		ptr=amudot.tashlum+roww;    requestedSeatsWorksheet[ptr].v=tmpTashlum[1];	
		
		ptr=amudot.tashlumPaid+roww;
		
		hasStillToPay=Number(tmpTashlum[1])-Number(delLeadingBlnks(requestedSeatsWorksheet[ptr].v));
		
		// if registration closed write request time
		ptr11=amudot.registrationClosedDateNTime+'2';  
	 tmp=delLeadingBlnks(requestedSeatsWorksheet[ptr11].v); 
	 ptr=amudot.requestDate+roww;
	 if(tmp){dd1= Date()}   //registration closed
	    else { dd1=' '};
		if(seatRelevantChangeRequest)requestedSeatsWorksheet[ptr].v=dd1;	
		
		afterClosingDate='$0';
	if ( delLeadingBlnks(requestedSeatsWorksheet[ptr].v) )afterClosingDate='$1';
	
	updateRowForNewSelection(roww);  // init all other field derived from selection string
	
	update_namesForSeat(roww);
	
	   
		
// write detailed request	
	xlsx.writeFile(workbook, XLSXfilename);
	
	
	ptr=amudot.email+roww;
	maill=delLeadingBlnks(requestedSeatsWorksheet[ptr].v);
	if (maill ){
	    subj='registraion request registered';
			txt='  ';
	    sendMail(maill,subj,txt);
	    }
 
	return true;						
}
function sortOrder(a,b){
  return Number(a.split('_')[0])-Number(b.split('_')[0]);
	}
// -----------------------------------------------------------------------------------------	
	
			
// delete leading blanks

 function delLeadingBlnks (str){
     str1=str;    if (! str1)return str1;
		 while (str1.substr(0,1)==' '){str1= str1.substr(1); 	  if (! str1.length)break;}		
		 while (str1.substr(str1.length-1,1)==' '){str1= str1.substr(0,str1.length-1); 	  if (! str1.length)break;}		

		 
		 return str1;
		 }

		 


// ----------------------------------------------------------------------------------------- 
ulamVlus=[2,3,1,1,0,0,0,0,0,0];
		
function setUlam( Minyan,ptr){

 requestedSeatsWorksheet[ptr].v  = ulamVlus[Minyan].toString();
 
 }
// ----------------------------------------------------------------------------------------- 
	
 var dbgOccupation=[]; 
  function setSeatOccupationLevel(holiday){     // holiday == 0 => both; holiday ==1 => rosh' holiday ==2 => kipur
	
	  seatOcuupationLevel.forEach(setToZero);; // clear previous values
	 for (ii=0; ii<lastSeatNumber+1;ii++)if (alreadyAssignedSeatsRosh[ii] || alreadyAssignedSeatsKipur[ii]){
	                 combinedAlreadyAssigned[ii]=true} else combinedAlreadyAssigned[ii]=false;
									 
		 for (member=firstSeatRow; member<lastSeatRow+1; member++){ 
		    sMember=member.toString();  
	if(dbgOccupation.indexOf(member) != -1){dbg=true}else dbg=false;			
				 toAssgnRoshMen=Number(requestedSeatsWorksheet[amudot.menRosh+sMember].v)-Number(requestedSeatsWorksheet[amudot.numberOfAssignedSeatsRoshMen+sMember].v);
				 toAssgnRoshMen=Math.max(toAssgnRoshMen,0);
 				 toAssgnRoshWomen=Number(requestedSeatsWorksheet[amudot.womenRosh+sMember].v)-Number(requestedSeatsWorksheet[amudot.numberOfAssignedSeatsRoshWomen+sMember].v);
				 toAssgnRoshWomen=Math.max(toAssgnRoshWomen,0);
 				 toAssgnKipurMen=Number(requestedSeatsWorksheet[amudot.menKipur+sMember].v)-Number(requestedSeatsWorksheet[amudot.numberOfAssignedSeatsKipurMen+sMember].v);
				 toAssgnKipurMen=Math.max(toAssgnKipurMen,0);
 				 toAssgnKipurWomen=Number(requestedSeatsWorksheet[amudot.womenKipur+sMember].v)-Number(requestedSeatsWorksheet[amudot.numberOfAssignedSeatsKipurWomen+sMember].v);
       	 toAssgnKipurWomen=Math.max(toAssgnKipurWomen,0);

		
				numSeatsMarkedForMen=requestedSeatsWorksheet[amudot.numberMarkedMen+sMember].v;
				numSeatsMArkedForWomen=requestedSeatsWorksheet[amudot.numberMarkedWomen+sMember].v;
				menRosh=Number(requestedSeatsWorksheet[amudot.menRosh+sMember].v);
				menKipur=Number(requestedSeatsWorksheet[amudot.menKipur+sMember].v);
				womenRosh=Number(requestedSeatsWorksheet[amudot.womenRosh+sMember].v);
				womenKipur=Number(requestedSeatsWorksheet[amudot.womenKipur+sMember].v);
				
				switch 	(holiday) {  
				     				 
				     case 0: 
						        requestedSeatsForMen=Math.max(toAssgnRoshMen,toAssgnKipurMen);
										requestedSeatsForWomen=Math.max(toAssgnRoshWomen,toAssgnKipurWomen); 
										seatStr=requestedSeatsWorksheet[amudot.notAssignedMarkedSeatsRosh+sMember].v; 
										markedButNoatAssignedList=seatStr.split('+');
										tempList=(requestedSeatsWorksheet[amudot.notAssignedMarkedSeatsKipur+sMember].v).split('+');
										  for (ik=0; ik<tempList.length;ik++)
											   if( markedButNoatAssignedList.indexOf( tempList[ik]) == -1)markedButNoatAssignedList.push(tempList[ik]);    
									
										
										break;
						 case 1: 
						        requestedSeatsForMen=toAssgnRoshMen; 
										requestedSeatsForWomen=toAssgnRoshWomen; 
										seatStr=requestedSeatsWorksheet[amudot.notAssignedMarkedSeatsRosh+sMember].v ; 
										markedButNoatAssignedList=seatStr.split('+');	
										break;	
			  	  	case 2: 
						        requestedSeatsForMen=toAssgnKipurMen;
										requestedSeatsForWomen=toAssgnKipurWomen;
										seatStr=requestedSeatsWorksheet[amudot.notAssignedMarkedSeatsKipur+sMember].v; 
										markedButNoatAssignedList=seatStr.split('+');	
										break;						
				                    }
										
					numSeatsMArkedForWomen=0;									
					for (ik=0; ik<markedButNoatAssignedList.length;ik++)if(isWoman[markedButNoatAssignedList[ik].split('_')[0]]==1)numSeatsMArkedForWomen++;	
					numSeatsMarkedForMen=markedButNoatAssignedList.length-	numSeatsMArkedForWomen;					
				  if( numSeatsMarkedForMen){ seatWeightForMen=Math.min(1, requestedSeatsForMen / numSeatsMarkedForMen)}
					   else seatWeightForMen=0; ;
					if( numSeatsMArkedForWomen){ seatWeightForWomen= Math.min(1,requestedSeatsForWomen / numSeatsMArkedForWomen)}
					  else seatWeightForWomen=0;
	
					
					seatStr=delLeadingBlnks(seatStr);
					if(!seatStr){continue;}
					seats=seatStr.split('+');
	        for (k=0; k<seats.length; k++){  
					  seat=Number(seats[k].split('_')[0]);
						combinedAlreadyAssigned[seat]=false;
						if( isWoman[seat]) {seatWeight=seatWeightForWomen; } else seatWeight=seatWeightForMen; 
						seatOcuupationLevel[seat]=seatOcuupationLevel[seat]+seatWeight; 

									 if(seatOcuupationLevel[seat] >9) seatOcuupationLevel[seat]=9;  // limit value to 9
	         };  
				
			
												 
																 
	      }   //end loop on members
				
			
				
		
				
				 return;
	
	}
// ----------------------------------------------------------------	
  
	function setToZero (element, index, array) {array[index]=0;}



   
    /*
     *  Initializes the sample application.
     */
    function initialize () {
        setupVariables();
        populateCache();
        setupTerminationHandlers();
				
                      };


    /*
     *  Start the server (starts up the sample application).
     */
    function start () {
        //  Start the app on the specific interface (and port).
      app.listen(port, ipaddress, function() {
            console.log('%s: Node server started on %s:%d ...',
                        (Date.now() ), ipaddress, port);
        });
    };

//  ----- handle member info request ------------------------

function  memberInfo(requestor,inputString){
	var respnArray=new Array;
	inp=inputString;
  inpData=inp.split('$'); 		
  name=delLeadingBlnks(inpData[0]);  holiday=['all','rosh','kipur'].indexOf(name);   
	if( (holiday !=-1 ) &&  ( requestor=='manage')   )     {     // send input for coloring the map
	         
						  setSeatOccupationLevel(holiday);
	            seatOccupationStr='aaa';
	             for (i=1; i<=lastSeatNumber; i++){
							 if (seatToRow[i] == 'NAN')continue;   
							   ocuup=seatOcuupationLevel[i];
								 ocuupTemp=ocuup;
							//	 if ( (! ocuup) && combinedAlreadyAssigned[i])ocuupTemp=1; 
	          //     if(ocuupTemp ) { 
						
								 ocupAdd=ocuup.toString().substr(0,6);
		 
		             ocupAdd=ocupAdd+'@'+namesForSeat[i];  //  get all names for seat to client
							   seatOccupationStr = seatOccupationStr +'+'+(i).toString()+'_'+ ocupAdd; 
	                //  }      //if occup
									
									}         // loop i
								
								 
							
					name=delLeadingBlnks(inpData[1]);  
					nameData='';
					if ( name) { 
					   rowNum= knownName(name)[0];
						 if(rowNum != -1){  
						       row=rowNum.toString();
									 ptr1=amudot.markedSeats+row; 
									 ptr2=amudot.assignedSeatsRosh+row; 
									 ptr3=amudot.assignedSeatsKipur+row; 
						       nameData=requestedSeatsWorksheet[ptr1].v;    // requested seats an assigned seats
						        nameData=nameData+'*'+requestedSeatsWorksheet[ptr2].v;  
										 nameData=nameData+'*'+requestedSeatsWorksheet[ptr3].v;         }
						
					          	}      
								respns=seatOccupationStr+'&'+nameData;
	       
								 }  // end of coloring 
	else	{					// regular name	  	
	
	 	 rawName=knownName(name);  
		 rowNum=rawName[0];
	
		 if (rowNum ==-1 ){  respns='---000'+rawName[1];}    // '�� �� ����'
	   else {  // recognized name
		     Row=rowNum.toString(); 
		     
		     mRosh=Number(requestedSeatsWorksheet[amudot.menRosh +Row].v);
				 wRosh=Number(requestedSeatsWorksheet[amudot.womenRosh +Row].v);
				 mKipur=Number(requestedSeatsWorksheet[amudot.menKipur +Row].v);
				 wKipur=Number(requestedSeatsWorksheet[amudot.womenKipur +Row].v); 
				 inputForMemberExists=true;
				 if ( (mRosh + wRosh + mKipur + wKipur) == 0) inputForMemberExists=false;      // member did not input a request; 
				 
		         
						 respnArray[positionInMsg.email]=requestedSeatsWorksheet[amudot.email+Row].v; 
						 respnArray[positionInMsg.addr]=requestedSeatsWorksheet[amudot.addr+Row].v;
						 respnArray[positionInMsg.phone]=requestedSeatsWorksheet[amudot.phone+Row].v;
						 respnArray[positionInMsg.gvarimRoshHashana]=requestedSeatsWorksheet[amudot.menRosh+Row].v;
						 respnArray[positionInMsg.gvarimKipur]=requestedSeatsWorksheet[amudot.menKipur+Row].v;
						 respnArray[positionInMsg.nashimRoshHashana]=requestedSeatsWorksheet[amudot.womenRosh+Row].v;
						 respnArray[positionInMsg.nashimKipur]=requestedSeatsWorksheet[amudot.womenKipur+Row].v;
						 respnArray[positionInMsg.minyanMuadafNashim]=(requestedSeatsWorksheet[amudot.preferedMinyanW+Row].v).substr(0,1);
						 if( ! inputForMemberExists) respnArray[positionInMsg.minyanMuadafNashim]='9';
						 respnArray[positionInMsg.esberNashim]=requestedSeatsWorksheet[amudot.preferedExplanationW+Row].v;
						 respnArray[positionInMsg.minyanMuadafGvarim]=(requestedSeatsWorksheet[amudot.preferedMinyanM+Row].v).substr(0,1);
						 if( ! inputForMemberExists) respnArray[positionInMsg.minyanMuadafGvarim]='9';
						 respnArray[positionInMsg.esberGvarim]=requestedSeatsWorksheet[amudot.preferedExplanationM+Row].v;
						 respnArray[positionInMsg.moreComments]=requestedSeatsWorksheet[amudot.cmnts+Row].v;
						 respnArray[positionInMsg.requestedSeats]=requestedSeatsWorksheet[amudot.markedSeats+Row].v;
						 respnArray[positionInMsg.requestDate]=requestedSeatsWorksheet[amudot.requestDate+Row].v;
						 respnArray[positionInMsg.assignedSeatsRosh]=requestedSeatsWorksheet[amudot.assignedSeatsRosh+Row].v;
						 respnArray[positionInMsg.assignedSeatsKipur]=requestedSeatsWorksheet[amudot.assignedSeatsKipur+Row].v;
						 respnArray[positionInMsg.tashlum]=requestedSeatsWorksheet[amudot.tashlum+Row].v;
						 respnArray[positionInMsg.tashlumPaid]=requestedSeatsWorksheet[amudot.tashlumPaid+Row].v;
						
						 respnArray[positionInMsg.stsfctnInFlr2YRSAgoYrWmn]=requestedSeatsWorksheet[amudot.stsfctnInFlr2YRSAgoYrWmn+Row].v;
						 respnArray[positionInMsg.stsfctnInFlr2YRSAgoYrMen]=requestedSeatsWorksheet[amudot.stsfctnInFlr2YRSAgoYrMen+Row].v;
						 respnArray[positionInMsg.TwoYRSAgoSeat]=requestedSeatsWorksheet[amudot.TwoYRSAgoSeat+Row].v;
						 
						 respnArray[positionInMsg.stsfctnInFlr3YRSAgoYrWmn]=requestedSeatsWorksheet[amudot.stsfctnInFlr3YRSAgoYrWmn+Row].v;
						 respnArray[positionInMsg.stsfctnInFlr3YRSAgoYrMen]=requestedSeatsWorksheet[amudot.stsfctnInFlr3YRSAgoYrMen+Row].v;
						 respnArray[positionInMsg.ThreeYRSAgoSeat]=requestedSeatsWorksheet[amudot.ThreeYRSAgoSeat+Row].v;
						 
						 respnArray[positionInMsg.memberShipStatus]=requestedSeatsWorksheet[amudot.memberShipStatus+Row].v;
						 
						 respnArray[positionInMsg.stsfctnInFlrLastYrWmn]=requestedSeatsWorksheet[amudot.stsfctnInFlrLastYrWmn+Row].v;
						 respnArray[positionInMsg.stsfctnInFlrLastYrMen]=requestedSeatsWorksheet[amudot.stsfctnInFlrLastYrMen+Row].v;
						 respnArray[positionInMsg.LastYrSeat]=requestedSeatsWorksheet[amudot.lstYrSeat+Row].v;
						 
						 respnArray[positionInMsg.issueInFloorWmn]=requestedSeatsWorksheet[amudot.issueInFloorWmn+Row].v;
						 respnArray[positionInMsg.issueinFloorMen]=requestedSeatsWorksheet[amudot.issueinFloorMen+Row].v;
						 respnArray[positionInMsg.issueBetweenFloors]=requestedSeatsWorksheet[amudot.issueBetweenFloors+Row].v;
						 respnArray[positionInMsg.nashimMuadaf]=requestedSeatsWorksheet[amudot.nashimMuadaf+Row].v;
						 respnArray[positionInMsg.gvarimMuadaf]=requestedSeatsWorksheet[amudot.gvarimMuadaf+Row].v;
						 
						 respnArray[positionInMsg.stsfctnInFlrThisYrWmn]=requestedSeatsWorksheet[amudot.stsfctnInFlrThisYrWmn+Row].v;
						 respnArray[positionInMsg.stsfctnInFlrThisYrMen]=requestedSeatsWorksheet[amudot.stsfctnInFlrThisYrMen+Row].v;
						 respnArray[positionInMsg.ThisYrSeat]=requestedSeatsWorksheet[amudot.ThisYRSSeat+Row].v;
						 
						 respns=respnArray.join('&');
			}		 
	
					    
							 
		 };  //end of 'regular name'
		return 	respns;
  
}
//

		 
//------------------------------------------------------------------------------------------------

/*
 *  main():  Main code.
 */


//   init variables
var familyNames= new Array();
var sortedFirstNames= new Array();
var hisName= new Array();
var herName= new Array();
sortedFirstNames= new Array();
minimumName= new Array();	
  
var reqMinyan = new Array;
var txtCodes = new Array;
var seatToRow = new Array;
var isWoman =new Array;
var startingRow;
var firstSeatRow;
var errorNumber;
var seats = new Array;
var incompatibilty;
var hasStillToPay;
var inputString;
var namesForSeat = new Array;
//closeRegistrationDate =new Date;
var afterClosingDate;
var alreadyAssignedSeatsRosh = new Array;
var alreadyAssignedSeatsKipur = new Array;
var combinedAlreadyAssigned= new Array;

var debugRequestsKeys = new Array;;
var debugRequestsAll = new Array;
var debugparam=''; 

var startingRowstsfction=6 ;


var  moedCode
var	 SidurUlam
var	 shlavBFilter;
var	 notPaidFilter;
var	 doneWithFilter;

var forgetList = new Array;

	var initDone=false;
	
var assgndBegunRosh=false;
var assgndBegunKipur=false;	
	
var 	maxCountSeats=[[0,0],[0,0]]; // rashi-martef  /  gvarim-nashim
var firstName = new Array;


var assignedPerUlam=[[[0,0],[0,0]],[[0,0],[0,0]]];  //[ulam][moed][gvarim-nashim]

var amudot ={name:'A',registrationClosedDateNTime:'C',requestDate:'D',permanentSeats:'E',zug_gever_yisha:'F',email:'G',addr:'H',phone:'I',
              menRosh:'J',menKipur:'K',womenRosh:'L',womenKipur:'M',preferedMinyanW:'N',
              preferedExplanationW:'O',preferedMinyanM:'P',preferedExplanationM:'Q',cmnts:'R',
							markedSeats:'S',numberMarkedMen:'T',numberMarkedWomen:'U',notAssignedMarkedSeatsRosh:'V',
							notAssignedMarkedSeatsKipur:'W',NumberOfNotAssignedMarkedSeatsMen:'X', NumberOfNotAssignedMarkedSeatsWomen:'Y',
							assignedSeatsRosh:'Z',assignedSeatsKipur:'AA',numberOfAssignedSeatsRoshMen:'AB',numberOfAssignedSeatsRoshWomen:'AC',
							numberOfAssignedSeatsKipurMen:'AD',numberOfAssignedSeatsKipurWomen:'AE',tashlum:'AF',tashlumPaid:'AG',
							stsfctnInFlrThisYrWmn:'AI',
							stsfctnInFlrThisYrMen:'AJ',ThisYRSSeat:'AK',
							stsfctnInFlrLastYrWmn:'AL',stsfctnInFlrLastYrMen:'AM',lstYrSeat:'AN',
							stsfctnInFlr2YRSAgoYrWmn:'AO',stsfctnInFlr2YRSAgoYrMen:'AP',TwoYRSAgoSeat:'AQ',stsfctnInFlr3YRSAgoYrWmn:'AR',
							stsfctnInFlr3YRSAgoYrMen:'AS',ThreeYRSAgoSeat:'AT',
							issueInFloorWmn:'AU',  issueinFloorMen:'AV',  issueBetweenFloors:'AW',   
							memberShipStatus:'AX',nashimMuadaf:'AY',gvarimMuadaf:'AZ',debugRequestsKeys:'BA',debugRequestsAll:'BB'
							};
	var lastCol='AZ';		
								
var amudotForMemberInfo	 ={name:'A',zug_gever_yisha:'F',email:'G',addr:'H',phone:'I',	memberShipStatus:'AU'};						
                       

											
	
amudotToClrInReqstdSeatsWhnGenNewYr= {registrationClosedDateNTime:'C',requestDate:'D',
              menRosh:'J',menKipur:'K',womenRosh:'L',womenKipur:'M',preferedMinyanW:'N',
              preferedExplanationW:'O',preferedMinyanM:'P',preferedExplanationM:'Q',cmnts:'R',
							markedSeats:'S',numberMarkedMen:'T',numberMarkedWomen:'U',notAssignedMarkedSeatsRosh:'V',
							notAssignedMarkedSeatsKipur:'W',NumberOfNotAssignedMarkedSeatsMen:'X', NumberOfNotAssignedMarkedSeatsWomen:'Y',
							assignedSeatsRosh:'Z',assignedSeatsKipur:'AA',numberOfAssignedSeatsRoshMen:'AB',numberOfAssignedSeatsRoshWomen:'AC',
							numberOfAssignedSeatsKipurMen:'AD',numberOfAssignedSeatsKipurWomen:'AE',tashlum:'AF',tashlumPaid:'AG',
							stsfctnInFlrLastYrWmn:'AI',stsfctnInFlrLastYrMen:'AJ',lstYrSeat:'AK',
							issueInFloorWmn:'AR',  issueinFloorMen:'AS',  issueBetweenFloors:'AT', 
							nashimMuadaf:'AV',gvarimMuadaf:'AW'}
									

		
var	positionInMsg={email:0,	addr:1,phone:2,	gvarimRoshHashana:3,gvarimKipur:4,nashimRoshHashana:5,nashimKipur:6,
		                minyanMuadafNashim:7, esberNashim:8,minyanMuadafGvarim:9,esberGvarim:10,moreComments:11,
									 requestedSeats:12,requestDate:13,assignedSeatsRosh:14,assignedSeatsKipur:15,tashlum:16,tashlumPaid:17,
									 stsfctnInFlr2YRSAgoYrWmn:18,stsfctnInFlr2YRSAgoYrMen:19,TwoYRSAgoSeat:20,stsfctnInFlr3YRSAgoYrWmn:21,
                   stsfctnInFlr3YRSAgoYrMen:22,ThreeYRSAgoSeat:23,memberShipStatus:24,stsfctnInFlrLastYrWmn:25,stsfctnInFlrLastYrMen:26,
									 LastYrSeat:27,issueInFloorWmn:28,  issueinFloorMen:29,  issueBetweenFloors:30,nashimMuadaf:31,gvarimMuadaf:32,
									 stsfctnInFlrThisYrWmn:33,stsfctnInFlrThisYrMen:34,ThisYrSeat:35
									 };
											  
	
var sortWeightsPtr={vetek:'F1',personalIssue:'F2',satisfactionHistory:'F3',satisfactionInFloor:'F4',horizontalDistance:'F5',
                    lastYearVS2YearsAgo:'F6',numberOfRequestedSeats:'F7',requestedSeatsPerFamilySize:'F8',Baby:'F10'}	
										

var amudotOfConfig={fromSeat:'A', toSeat:"B",reltvRowQual:'C',open_badSeats:'D',ezor:'E',ulam:'F',X_forSlantedRow:'H',Y_forSlantedRow:'I'};		

var amudotForStsfctn=[	amudot.stsfctnInFlrThisYrWmn,amudot.stsfctnInFlrThisYrMen,amudot.ThisYRSSeat,
							          amudot.stsfctnInFlrLastYrWmn,amudot.stsfctnInFlrLastYrMen,amudot.lstYrSeat,
							          amudot.stsfctnInFlr2YRSAgoYrWmn,amudot.stsfctnInFlr2YRSAgoYrMen,amudot.TwoYRSAgoSeat,
												amudot.stsfctnInFlr3YRSAgoYrWmn,amudot.stsfctnInFlr3YRSAgoYrMen,amudot.ThreeYRSAgoSeat];	
												
var hebrewLetters={};
/*alef:0, bet:0, gimel:0, dalet:0, hei:0, vav:0, zain:0, chet:0, tet:0, yod:0, khaf_sofit:0,khaf:0,
        lamed:0, mem_sofit:0, mem:0,nun_sofit:0, nun:0, samech:0, ayin:0, pe_sofit:0, pe:0, tzadi_sofit:0, tzadi:0,
				kuf:0, reish:0, shin:0, taf:0};				*/														
		
var seatOcuupationLevel = new Array;			
var requestedSeatsWorksheet ;
var passwordsWS;
var	mngmntPASSW;
var	gizbarPASSW;	
var moshavimPASSW;
var errCodeWS;	
var seatToRowWS;
var shulConfigerationWS;
//var nashimGvarimRegions=[];
var UlamMartef=[];
var ezorOfSeat=[];
var configRowOfSeat=[];
var sortWeightsSheet;	

var lastYearInit='-';

var firstConfigRow;

var vetekWeight;
var personalIssueWeight;
var satisfactionHistoryWeight;
var satisfactionInFloorWeight;
var horizontalDistanceWeight;
var lastYearVS2YearsAgoWeight;
var numberOfRequestedSeatsWeight;
var requestedSeatsPerFamilySizeWeight;
var BabyWeight;
var stsfctionColction=[];
var assignedSeatslist=[[[],[]],[[],[]]];

var originalReqSeats=[];
var originalReqSeatPriority=[];
var priorityFactorConst=0.95;

var transactionHistoryFirstRow=250;

var dbgStsfction=false;
var aChagRslts=[];									
rqstdSeats=[[],[]];  // men list and women list
var rqstdRows=[[],[]];
var assgndRows= [ [[],[]], [[],[]] ];

var badSeats=[];


var shiras_backup=false;			
									
 app=express();
 
 app.use(function(req, res, next) {
  res.header("Access-Control-Allow-Origin", "*");
  res.header("Access-Control-Allow-Headers", "Origin, X-Requested-With, Content-Type, Accept");
  next();
});

initialize();

         /*     init new files in debug  */

	localFileDir='/data/';    
	//localFileDir='D_';
		 
XLSXfilename=localFileDir	+'membersRequests.xlsx';  
EmptyXLSXfilename=	localFileDir+'EmptymembersRequests.xlsx';           
seatsOrderedFileName=	localFileDir+'seatsOrdered.xlsx';
errPasswFilename=localFileDir+'empty.xlsx';
supportTblsFilename=	localFileDir+'supportTables.xlsx';  
registeredListFilename=localFileDir+'registeredList.xlsx';
BackupFilename= localFileDir+'BackupMembersRequests.xlsx';     


	
 /*debug code   1 
 
 

tmpfile=fs.readFileSync('supportTables.xlsx');  
console.log('2');                      
	fs.writeFileSync(supportTblsFilename, tmpfile);
	console.log('3');
supportWB=xlsx.readFile(supportTblsFilename);
  //	    supportWB=xlsx.readFile('supportTables.xlsx');  
console.log('4');
tmpfile=fs.readFileSync('membersRequests.xlsx');
console.log('5');
	fs.writeFileSync(XLSXfilename, tmpfile);
	workbook = xlsx.readFile(XLSXfilename);
	requestedSeatsWorksheet = workbook.Sheets['HTMLRequests'];  
	
	requestedSeatsWorksheet[amudot.debugRequests].v=' '; // on loading  anew database all debugs are reset
	
	*/ //////// - end debug code  1  
	
	tmpfile=fs.readFileSync('supportTables.xlsx');  
                    
	fs.writeFileSync(supportTblsFilename, tmpfile); 

	supportWB=xlsx.readFile(supportTblsFilename); 
	
	passwordsWS=supportWB.Sheets['passwords'];
	mngmntPASSW=passwordsWS['B1'].v;  
	gizbarPASSW=	passwordsWS['B2'].v;	
	debugPASSW=	passwordsWS['B3'].v;	
  moshavimPASSW=passwordsWS['B4'].v;
	
	
errCodeWS=supportWB.Sheets['errorCodes'];
for (i=1; i<50; i++){
 ptr1='A'+(i).toString();
 if (errCodeWS[ptr1].v == '$$$') break;
 ptr2='B'+(i).toString();
 txtCodes[i]=errCodeWS[ptr2].v;    
      };
	
	
	
	//   check if file exists try the code next project
	
	var 	numberOfDB_loadTrials=1;

	checkIf_memberRequstsExist();  // if exists finish init process
	


	function checkIf_memberRequstsExist(){console.log('entered ckeckif');
	
	if(initDone)return;
	
   
	fs.access(XLSXfilename, error => {
    if (!error) {
        // The check succeeded
				
		
	      console.log('DB found after '+numberOfDB_loadTrials+' trials');
		    initCompletion();
    } else {
        // The check failed
				console.error('DB not yet found after '+numberOfDB_loadTrials+' trials');
	
	      numberOfDB_loadTrials++;
	    setTimeout(checkIf_memberRequstsExist, 6000);	//check every 1 minutes
    }
}); 
	        
					
} // function
	
	
	
	
	
	
	
	

//-----------------------init gmail ------------------------------------------
   

    var transporter = nodemailer.createTransport({
        service: 'Hotmail',
        auth: {
            user: 'kehilatarielseats@outlook.com', // Your email id
            pass: 'kehila11' // Your password
        }
    });
		

             
// -----------------------------------------------------------------------

var workbook, supportWB;
var dayOfLastBackup=-1; 

function initCompletion(){

 workbook = xlsx.readFile(XLSXfilename);  

 

//read error codes  from supportTables.xlsx            





var debugRows=[];

		
		
initFromFiles(''); // init info from files for last year
		
	checkDoubeeAssignments();

sortDB();
	


setTimeout(backupRequests, 60000);	//check every 10 minutes

lastCol='AZ'; 
var numOfColsInNewSheet=colNametoNumber(lastCol)+10;  // 10 is spare
var numOfRowsInNewSheet=lastSeatRow+40;  // 40 spare for new names

console.log('init done at  '+getPrintableDate(0)+'   DST ignored');  // [2] is date + time

initDone=true;
} // end of initCompletion
//------------------------------------------------------------

function initFromFiles(yearToInitFrom){
 //  if(yearToInitFrom == lastYearInit)return;
	// lastYearInit=yearToInitFrom;
	 
	 
   initValuesOutOfSupportTablesXLSX_file (yearToInitFrom);
	
    initValuesOutOfHtmlRequestsXLSX_file(yearToInitFrom);   //init values
		
		
		}

//------------------------------------------------------------
  
	function checkDoubeeAssignments(){
	   var doubles=[];
		 doubles_idx=0;
	   assignedCol=amudot.assignedSeatsRosh;
	   for (moed=1; moed<3;moed++){
		  moedstr=moed.toString();
	     for (i=firstSeatRow;i<lastSeatRow;i++){ 
         row1=i.toString();
	       assigned_row1_STR=delLeadingBlnks(requestedSeatsWorksheet[assignedCol+row1].v);
				 if( ! assigned_row1_STR) continue;
				 assigned_row1=assigned_row1_STR.split('+');
				 for (j=i+1; j<lastSeatRow+1;j++){
				    row2=j.toString();
						assigned_row2_STR=delLeadingBlnks(requestedSeatsWorksheet[assignedCol+row2].v);
					  if( ! assigned_row2_STR) continue;

				    assigned_row2=assigned_row2_STR.split('+');
						for (k=0; k<assigned_row2.length;k++){
						  if (assigned_row1.indexOf(assigned_row2[k])  != -1){
							                      doubles[doubles_idx]=row1+'-'+row2+'- moed='+moedstr;
																		doubles_idx++;
																		
																		}
										} // for k								
	}  //j
	}  // i
	       assignedCol=amudot.assignedSeatsKipur;
	} //moed
	
	  if ( ! doubles_idx)return; // no doubles
		  console.log('doubles');
			console.log(doubles);
		
		
	} // function


	
//-------------------------------------------------------------------
 function checkRequestedAssgnmntForDoubles(reqstedAssgmnt, moed,checkedRow){  // function returns true if a problem IS found
 
 var assignedCol,i,j,tmp,row,assignmentForRow_STR;
 assnmentsToVerify=reqstedAssgmnt.split('+');
 
 if(moed==1){assignedCol=amudot.assignedSeatsRosh}else assignedCol=amudot.assignedSeatsKipur;
  for (i=firstSeatRow;i<lastSeatRow+1;i++){ 
         if (i == checkedRow)continue;
				 row=i.toString();
	       assignmentForRow_STR=delLeadingBlnks(requestedSeatsWorksheet[assignedCol+row].v);
				 if( ! assignmentForRow_STR) continue;
				 assignmentForRow=assignmentForRow_STR.split('+');
         for(j=0;j<assignmentForRow.length;j++)if ( assnmentsToVerify.indexOf(assignmentForRow[j]) != -1)return true;
				 } // i
	return false; // no problem found			 
				      

}
//-------------------------------------------------------------------

function colNametoNumber(col){
   var alphabet='ABCDEFGHIJKLMNOPQRSTUVWXYZ';
    if (col.length ==1) return alphabet.indexOf(col)+1;
		col1=col.substr(0,1);  col2=col.substr(1,1);
		num= 26*(alphabet.indexOf(col1)+1)+alphabet.indexOf(col2)+1;
		
		return num;
		       
	}
//-------------------------------------------------------------
function backupRequests(){
    var d1 = new Date();  
    var hour_Greenwich_Mean_Time = Number(d1.getHours());
 
    weekDay=d1.getDay();  

if (   ! shiras_backup){		// do the following only for scheduled backups
// handle forgetList

     for (iikk=firstSeatRow; iikk< lastSeatRow+1;iikk++){
    		 if (forgetList[iikk]){
				      countr=Number(forgetList[iikk].split('$')[1]);
							if (countr>0)countr--;              forgetList[iikk]= forgetList[iikk].split('$')[0]+'$'+countr.toString();
							if( ! countr )forgetList[iikk]='';
							}
				}
			
				
// end of forget list handling

							 
		
		 if( weekDay == dayOfLastBackup) {    
		    setTimeout(backupRequests, 600000);	//check every 10 minutes
		    return;
				}
	} // if  ! shiras_backup
		
    if( (hour_Greenwich_Mean_Time == 0 ) || ((shiras_backup) && (initDone))){      // once a day; at night; when value=0 => in the winter 2am; in summer 3am// ! hour_Greenwich_Mean_Time
		
    	 xlsx.writeFile(workbook, BackupFilename);
	 
	     dayOfLastBackup=weekDay;
	     var mailOptions = {
            from: 'kehilatarielseats@outlook.com', // sender address
            to: 'kehilatarielseats@gmail.com', // list of receivers
            subject: 'backupCreated', // Subject line
            text: 'backup',  // plaintext body
            attachments: [
                {   // file on disk as an attachment
               filename: 'requestsBackup.xlsx',
               path: BackupFilename // stream this file
              }  ]                  
											
					};
								    
     transporter.sendMail(mailOptions, function(error, info){
         if(error)  console.log('send backup mail reported an error=='+error+'  at '+d1);
	    })
	 console.log('backup created at '+d1);
	 
	 }
if (! shiras_backup)setTimeout(backupRequests, 600000);	//check every 10 minutes

shiras_backup=false;

}	 


//-------------------------------------------------------------
 function sendMail(addr,subj,txt){
      
      var mailOptions = {
            from: 'kehilatarielseats@outlook.com', // sender address
            to: addr, // list of receivers
            subject: subj, // Subject line
            text: txt,  // plaintext body
  	      	};
									    
     transporter.sendMail(mailOptions, function(error, info){
         if(error)  console.log('send  mail to '+addr+' subj='+subj+' text='+txt+'  info='+info+' reported an error=='+error);
	    })
 
 
 }

//--------------------------------------------------------------- 
function initValuesOutOfSupportTablesXLSX_file(yearToInitFrom){  
 var tmp;
 var badList=[];
 
//read Seat to Row from supportTables.xlsx  
//console.log('yearToInitFrom='+yearToInitFrom);
   maxCountSeats=[[0,0],[0,0]]; // rashi-martef  /  gvarim-nashim
		
	lastSeatNumber=0;    
	seatToRowWS=supportWB.Sheets['seatToRow'+yearToInitFrom]; 
  shulConfigerationWS=supportWB.Sheets['shulConfigeration'+yearToInitFrom];  
	
	for (i=0;i<1500;i++){  
	      isWoman[i]='';
				configRowOfSeat[i]=''; 
				}
 
 badSeats=[];
 
 for (firstConfigRow=1;firstConfigRow<20;firstConfigRow++){
         ptr=amudotOfConfig.fromSeat+firstConfigRow.toString(); 
				 tmp=shulConfigerationWS[ptr].v;    
				 if ( ! isNaN(tmp) )break;
			}
	if (firstConfigRow ==20){console.log('error in supporttables'); return}	 
 
 
  for (i=firstConfigRow; i<1500; i++){
	  row=(i).toString();
    ptrFrom=amudotOfConfig.fromSeat+row;  
		vlu=shulConfigerationWS[ptrFrom].v;
		if (vlu == '$$$')break;
		
		
		Stfrom=Number(vlu);
		
		if (! Stfrom){  // if from seat = 0 then info at "edge seats" is a list of "bad seats"
		   ptr=amudotOfConfig.open_badSeats+row;
		   badList=(shulConfigerationWS[ptr].v).split('+');
			 for (ii=0; ii< badList.length; ii++)badSeats.push(Number(badList[ii]));  // add this row bad seats to global list of bad seats
		  continue;
			}
		ptrTo=amudotOfConfig.toSeat+row;
		StTo=Number(shulConfigerationWS[ptrTo].v);
		
		ptrEzor=amudotOfConfig.ezor+row;
		ezor=shulConfigerationWS[ptrEzor].v;
		
		ptrUlam=amudotOfConfig.ulam+row;
		ulam=shulConfigerationWS[ptrUlam].v;
		if (ulam.substr(0,1) != 'n'){nashim=0;} else nashim=1;  // gender value for nashim in maxCountSeats is 1 
		
		itmp=ulam.indexOf(' ');
		tmp=ulam.substr(itmp+1,1);
		UlamOrMartef='m';  ind1=1;
				if (tmp != 'm') {  UlamOrMartef='u';  ind1=0;};
		
		maxCountSeats[ind1][nashim] += StTo-Stfrom+1;
		  
		for(st=Stfrom; st<StTo+1; st++){
		    isWoman[st]=nashim;
		    UlamMartef[st]=UlamOrMartef;
				ezorOfSeat[st]=ezor;
				configRowOfSeat[st]=i;
				
			}	
	}			
		
		
		i=1;
		ptr1=amudotOfConfig.fromSeat+(i).toString();
		nextSt=seatToRowWS[ptr1].v	;
		while (nextSt != '$$$'){
   
      alreadyAssignedSeatsRosh[i]=' '; 
      alreadyAssignedSeatsKipur[i]=' ';
      seatToRow[i]=seatToRowWS[ptr1].v;
  
	
		i++;  if(i > 1500){console.log('error in seat to row'); return;};
		ptr1=amudotOfConfig.fromSeat+(i).toString();
		nextSt=seatToRowWS[ptr1].v	;
   };
		
		lastSeatNumber=i-1;
			
	
 // get  sortWeights
sortWeightsSheet=supportWB.Sheets['sortWeights'];   

vetekWeight=Number(sortWeightsSheet[sortWeightsPtr.vetek].v);
personalIssueWeight=Number(sortWeightsSheet[sortWeightsPtr.personalIssue].v);
satisfactionHistoryWeight=Number(sortWeightsSheet[sortWeightsPtr.satisfactionHistory].v);
satisfactionInFloorWeight=Number(sortWeightsSheet[sortWeightsPtr.satisfactionInFloor].v);
horizontalDistanceWeight=Number(sortWeightsSheet[sortWeightsPtr.horizontalDistance].v);
lastYearVS2YearsAgoWeight=Number(sortWeightsSheet[sortWeightsPtr.lastYearVS2YearsAgo].v);
numberOfRequestedSeatsWeight=Number(sortWeightsSheet[sortWeightsPtr.numberOfRequestedSeats].v);
requestedSeatsPerFamilySizeWeight=Number(sortWeightsSheet[sortWeightsPtr.requestedSeatsPerFamilySize].v);
BabyWeight=Number(sortWeightsSheet[sortWeightsPtr.Baby].v);

// upload hebrew letters
HebrewLettersSheet=supportWB.Sheets['hebrewletters'];   
for (i=1; i<28;i++){
  row=i.toString();
	tmp1=HebrewLettersSheet['A'+row].v; 
	tmp2=HebrewLettersSheet['B'+row].v;  
	
  hebrewLetters[tmp2]=tmp1;
	
	}
	//console.log(hebrewLetters);
	
}
//-------------------------------------------------------------
function initValuesOutOfHtmlRequestsXLSX_file(yearToInitFrom){
var i,j,k,l, tmp,tmp1,tmp2,ptrA,ptrB,famParts, fullName, sameFamName,   minNameSet;
tmp=new Array;
indices=new Array;
for(i=1; i<lastSeatNumber+1; i++){
      seatOcuupationLevel[i]=0;    // clear and set array size 
			namesForSeat[i]='$/';
			};

requestedSeatsWorksheet = workbook.Sheets['HTMLRequests'+yearToInitFrom];
				 			
	// reload debug requests		           
 // debugRequestsWorkSheet=workbook.Sheets['debugRequests'];
	
 debugRequestsKeys-[];
 debugRequestsAll=[];
 k=0; 
	for(i=0;i<20;i++){
	   ptrA=amudot.debugRequestsKeys+(i+3).toString(); // first row is 3
	   tmp=requestedSeatsWorksheet[ptrA].v;
		 if (tmp == '$$$')continue;
		 debugRequestsKeys[k]=tmp;
		 ptrB=amudot.debugRequestsAll+(i+3).toString(); 
	   debugRequestsAll[k]=requestedSeatsWorksheet[ptrB].v;
		 k++;
		
		 };
	//console.log('initvalues length='+debugRequestsKeys.length);
 
 
 
	for(i=0;i<200;i++){    // clear name tables
	       familyNames [i]='';
				 hisName[i]='';
				 herName[i]='';
				 minimumName[i]='';
				 };  
	
	
	// firstName=[];
	 	 firstSeatRow=0;
	 for (i=2;i<200;i++){ 
	 row=i.toString();
	 pointerCell=amudot.name+row;

	 
	 cell=requestedSeatsWorksheet[pointerCell]; 
	  if(! cell) continue;
		cell.v=delLeadingBlnks(cell.v);
		if ( ! cell.v )continue;
		if ( firstSeatRow == 0) firstSeatRow=i;   // first name  row
		
	     fullName=cell.v; 
			 if (fullName == '$$$'){ lastSeatRow=i-1; break;}
			 famParts= fullName.split('*');
			 cell.v=famParts.join('*'); // remove un necessary blanks
			 familyNames [i-firstSeatRow]=famParts[0];
				 hisName[i-firstSeatRow]=famParts[1];
				 herName[i-firstSeatRow]=famParts[2];
				 /*
			 if (famName.substr(famName.length-1,1)=='*') {
			     famName=famName.substr(0,famName.length-1);
					 theFirstName=famParts[famParts.length-1];
					 theFirstName=theFirstName.substr(0,theFirstName.length-1);
					 firstName[i-firstSeatRow]=theFirstName;
			     } 
					 else firstName[i-firstSeatRow]='';
					 
	     familyNames[i-firstSeatRow]= famName; 
		   */
			 forgetList[i]='';
			  
		   mRosh=Number(requestedSeatsWorksheet[amudot.menRosh +row].v);
				 wRosh=Number(requestedSeatsWorksheet[amudot.womenRosh +row].v);
				 mKipur=Number(requestedSeatsWorksheet[amudot.menKipur +row].v);
				 wKipur=Number(requestedSeatsWorksheet[amudot.womenKipur +row].v); 
				
				 if ( (mRosh + wRosh + mKipur + wKipur) == 0) continue; // member did not input a request
	
				 closeSeats(1,i);
				closeSeats(2,i);  
						 
	   }  
		 
		
		familyNames.length=i-firstSeatRow;
		
	
	 if (i>190)reportAnError('no $$$ at end of family names'); 
	 
	 // set minimum names
	   
		
			for(i=0; i< familyNames.length;i++){ 
			indices=[i]; 
			  if(  minimumName[i] ) continue;  // minimum name already set
	//      minimumName[i]=familyNames[i];               console.log('i='+i+' minimumName[i]='+minimumName[i]+' familyNames[i]='+familyNames[i] );
		
	      for (j=i+1; j<familyNames.length;j++)if (familyNames[i] == familyNames[j] )indices.push(j);
				
				if(  indices.length ==1 ){minimumName[i]=familyNames[i];   continue; } // this fam name appears only once
				
				for (k=0; k<indices.length;k++){
				   l=indices[k];
				  if (! hisName[l]){minimumName[l]=familyNames[i]+' '+herName[l]; continue}; // no man
					
				  for(m=0; m<indices.length ;m++){
					  n=indices[m];
					  if (l !=n) if(hisName[l] == hisName[n]){minimumName[l]=familyNames[i]+' '+herName[l];	break};
	           } //for m
						 if ( ! minimumName[l])minimumName[l]=familyNames[i]+' '+hisName[l];
				} // for k
				
				
				
				
		} // for i
		
	
		
		// set sortedFirstNames
		 for (i=0; i<	familyNames.length;i++) if ( hisName[i] <herName[i]){ sortedFirstNames[i]=hisName[i]+'*'+herName[i]}
		                                                                 else sortedFirstNames[i]=herName[i]+'*'+hisName[i] ;
																																		 
	 
	 for (i=1; i<lastSeatNumber+1; i++){
               alreadyAssignedSeatsRosh[i]=' '; 
               alreadyAssignedSeatsKipur[i]=' '; 
							 }
	
//return;
	init_notAssigenedMarked('rosh');
	  init_notAssigenedMarked('kipur');

    CountAssignedPerMoed_PerUlam();
	
	}
	
 
// -----------------------------------------------------------------------
function counSeatsInEzor(row,moed,ezor){
var reslt = new Array;



        menRosh=Number(requestedSeatsWorksheet[amudot.menRosh+row].v);
				menKipur=Number(requestedSeatsWorksheet[amudot.menKipur+row].v);
				womenRosh=Number(requestedSeatsWorksheet[amudot.womenRosh+row].v);
				womenKipur=Number(requestedSeatsWorksheet[amudot.womenKipur+row].v);
				
				switch 	(moed) {  
				     				 
				     case 3: 
						        reslt[1]=Math.max(menRosh,menKipur);
										reslt[0]=Math.max(womenRosh,womenKipur); 
										
										
										break;
						 case 1: 
						        reslt[1]=menRosh; 
										reslt[0]=womenRosh; 
										
										break;	
			  	  	case 2: 
						        reslt[1]=menKipur;
										reslt[0]=womenKipur;
										
										break;						
				                    }

     if(ezor != '3'){  // not 'all' but filtered
        if (requestedSeatsWorksheet[amudot.nashimMuadaf+row].v != ezor)reslt[0]=0;	
				if (requestedSeatsWorksheet[amudot.gvarimMuadaf+row].v != ezor)reslt[1]=0;	
          }
		 return reslt;
}

//------------------------------------------------------------------------------------ 
function simplifyName(name){
var tmp,simplified;
tmp=name.split('*');
if  ( (! tmp[1])  && ( !tmp[2]) ) return tmp[0];
 if (tmp[1]  && tmp[2] ){simplified=tmp[0]+' '+tmp[1]+' '+hebrewLetters.vav+tmp[2] } else simplified=tmp[0]+' '+tmp[1]+tmp[2];
	return simplified;	
	
	}
	  

//------------------------------------------------------------------------------------ 


function update_namesForSeat(row){
debug1=false;     //if (row  =='173') debug1=true; 
  nm=simplifyName(delLeadingBlnks( requestedSeatsWorksheet[amudot.name +row].v));
  
 
	markedSeatsSTR=delLeadingBlnks( requestedSeatsWorksheet[amudot.markedSeats +row].v); 
	if(markedSeatsSTR){ // seats were marked
	    markedSeats=markedSeatsSTR.split('+');
			for (il=0; il<markedSeats.length;il++){
			   aSeat=Number(markedSeats[il].split('_')[0]);  
				 namesForSeatParts=namesForSeat[aSeat].split('/');
				 namesTemp=delLeadingBlnks(namesForSeatParts[1]); 
         if(namesTemp){namesToAttach=namesTemp.split('$')} else namesToAttach=[]; 
				 if (namesToAttach.indexOf(nm) == -1){
				       namesToAttach.push(nm); 
							 namesForSeatParts[1]=namesToAttach.join('$');
							 namesForSeat[aSeat]=namesForSeatParts.join('/'); 
							 }
         }
		}
		
	assgnedForRoshSTR=delLeadingBlnks( requestedSeatsWorksheet[amudot.assignedSeatsRosh +row].v);	
	if(assgnedForRoshSTR){
	  assgnedForRosh=assgnedForRoshSTR.split('+');
		for (il=0; il<assgnedForRosh.length;il++){
			   aSeat=Number(assgnedForRosh[il]);
				 namesForSeatParts=namesForSeat[aSeat].split('/');
         namesAssigned=namesForSeatParts[0].split('$');
	       namesAssigned[0]=nm;   
				 namesForSeatParts[0]=namesAssigned.join('$'); 
	       namesForSeat[aSeat]=namesForSeatParts.join('/');
			
		}		
	};
	
	assgnedForKipurSTR=delLeadingBlnks( requestedSeatsWorksheet[amudot.assignedSeatsKipur +row].v);	
	if(assgnedForKipurSTR){
	  assgnedForKipur=assgnedForKipurSTR.split('+');
		for (il=0; il<assgnedForKipur.length;il++){
			   aSeat=Number(assgnedForKipur[il]);
				 namesForSeatParts=namesForSeat[aSeat].split('/');
         namesAssigned=namesForSeatParts[0].split('$');
	       namesAssigned[1]=nm;
				 namesForSeatParts[0]=namesAssigned.join('$');
	       namesForSeat[aSeat]=namesForSeatParts.join('/');
		}		
	};			 

}
//-------------------------------------------------------------------------- 
 function filterAndSort(){
 var tmp;
 var counts=[];
 var listToSend=[];
 var tempList=[];
 idx=0;
 var indxToMinNames=[];
 moed=['rosh','kipur','all'].indexOf(moedCode)+1;
 // start with filtering

 for (i=firstSeatRow;i<lastSeatRow+1;i++){ 
   row=i.toString();
	 
   if(shlavBFilter=='true'){       
	  AfterClosingDate=delLeadingBlnks(requestedSeatsWorksheet[amudot.registrationClosedDateNTime +row].v); 
	  if(AfterClosingDate)continue;
		} //shlavBFilter
		
		tmpVl=Number(requestedSeatsWorksheet[amudot.menRosh+row].v)+Number(requestedSeatsWorksheet[amudot.womenRosh+row].v)
		      +Number(requestedSeatsWorksheet[amudot.menKipur+row].v)+Number(requestedSeatsWorksheet[amudot.womenKipur+row].v);
		if ( !	tmpVl ) continue;  // no request made		

		if (notPaidFilter=='true'){
		    paidValue=delLeadingBlnks(requestedSeatsWorksheet[amudot.tashlumPaid +row].v); 
	      if( ! paidValue)continue;
		}  // notPaidFilter
		
		   toAssgnRoshMen=Number(requestedSeatsWorksheet[amudot.menRosh+row].v)-Number(requestedSeatsWorksheet[amudot.numberOfAssignedSeatsRoshMen+row].v);
 			 toAssgnRoshWomen=Number(requestedSeatsWorksheet[amudot.womenRosh+row].v)-Number(requestedSeatsWorksheet[amudot.numberOfAssignedSeatsRoshWomen+row].v);
 			 toAssgnKipurMen=Number(requestedSeatsWorksheet[amudot.menKipur+row].v)-Number(requestedSeatsWorksheet[amudot.numberOfAssignedSeatsKipurMen+row].v);
 			 toAssgnKipurWomen=Number(requestedSeatsWorksheet[amudot.womenKipur+row].v)-Number(requestedSeatsWorksheet[amudot.numberOfAssignedSeatsKipurWomen+row].v);
		
		
		
		// name not filtered
		
		counts=counSeatsInEzor(row,moed,SidurUlam);  
		if ( ! (counts[0]+ counts[1]) ) continue;    // no seat requested for current ulam
		
		nameToKeep= delLeadingBlnks(requestedSeatsWorksheet[amudot.name +row].v);   //.split('*');
		// if (nameToKeep[1]  && nameToKeep[2] ){nameToKeep=nameToKeep[0]+' '+nameToKeep[1]+' '+hebrewLetters.vav+nameToKeep[2] } else nameToKeep=nameToKeep[0]+' '+nameToKeep[1]+nameToKeep[2];
		
		tempList[idx]=nameToKeep+'$'+calcSortParam(row)[0]+'$'+calcSortParam(row)[1]+
		                '$'+counts[0].toString()+'$'+counts[1].toString()
										+'$'+toAssgnRoshMen.toString()+'$'+toAssgnRoshWomen.toString()+'$'+toAssgnKipurMen.toString()+'$'+toAssgnKipurWomen.toString();
		
		//console.log(tempList[idx]);
		idx++
   }// for


 
 // sort step  A
 tempList=tempList.sort(sortOrderFirstParam);

 //move to tempList_2 upto amount of available seats
 countWomen=0;
 countMen=0;
 tempList_2=[];
 idx=0;
 maxCountMen=1000;  // for the case of sidurUlam=3 which means no cutting of the list
 maxCountWomen=1000;
 
 //[ulam][gvarim-nashim]   
 
 if(SidurUlam=='1'){maxCountMen=assignedPerUlam[0][1]; maxCountWomen= assignedPerUlam[0][0]};
 if(SidurUlam=='2'){maxCountMen=assignedPerUlam[1][1]; maxCountWomen= assignedPerUlam[1][0]};

 for (i=0;i<tempList.length;i++){

   countWomen=countWomen+Number(tempList[i].split('$')[3]);
   countMen=countMen+Number(tempList[i].split('$')[4]);
   if( (countMen>maxCountMen) || (countWomen>maxCountWomen) ) break;
	 tempList_2[i]=tempList[i];
	} 
	 
 //  sort step  B
   tempList_2=tempList_2.sort(sortOrderSecondParam);
	
  

 // second filtering stage of those that are already done and stripping of not required info
   vlu=[];
	 idx=0;
	 
	 for(i=0;i<tempList_2.length;i++){
	 
         vlu=tempList_2[i].split('$');
		     toAssgnRoshMen=Number(vlu[5]);
 				 toAssgnRoshWomen=Number(vlu[6]);
 				 toAssgnKipurMen=Number(vlu[7]);
 				 toAssgnKipurWomen=Number(vlu[8]);
         doneWithFlag=false;  

				 switch 	(moedCode) {  
				     	case 'rosh':  
							 if ( ( ! toAssgnRoshMen) && ( ! toAssgnRoshWomen) )doneWithFlag=true;
							 break;
		          case 'kipur':  
							 if ( ( ! toAssgnKipurMen) && ( ! toAssgnKipurWomen) )doneWithFlag=true;
							 break;
							case 'all':  
							 if ( ( ! toAssgnRoshMen) && ( ! toAssgnRoshWomen) && ( ! toAssgnKipurMen) && ( ! toAssgnKipurWomen) )doneWithFlag=true;
							 break;
		      }  // switch
					
				notPartOfTheListFlag='+';	
		    if ( (doneWithFilter=='true') && doneWithFlag)notPartOfTheListFlag='-';
				listToSend[idx]=notPartOfTheListFlag+vlu[0]; // strip not required info
				indxToMinNames[idx]=(knownName(simplifyName(vlu[0]))[0]-firstSeatRow).toString();
				idx++; 
		}  //doneWithFilter
		

 strr=listToSend.join('$')+'@'+indxToMinNames.join('$');   
 return strr;
 }
 

//--------------------------------------------------------------------------     
function calcSortParam(row){
 var calcResult=[];
 var tmp;
 var  tmp1;
 var tmp2, tmp3;;
 //   calc family sort calue
 
// part personal problem family between floors
part1= personalIssueWeight*Number(requestedSeatsWorksheet[amudot.issueBetweenFloors+row].v);

// part vetek
part2=vetekWeight*Number(requestedSeatsWorksheet[amudot.memberShipStatus+row].v);

//part mishkal koma in the past
//MakomMinus3Yrs=Number(requestedSeatsWorksheet[amudot.ThreeYRSAgoSeat+row].v);
tmp3=requestedSeatsWorksheet[amudot.ThreeYRSAgoSeat+row].v;
if (isNaN(tmp3)){ 
         tmp=tmp3.split('*');
         if ( !tmp[1]){tmp1=Number(tmp[0])} else tmp1=Number(tmp[1]);}
			else tmp1=tmp;	
			if( ! delLeadingBlnks(tmp3) )tmp1=0; 
MakomMinus3Yrs=tmp1;

//MakomMinus2Yrs=lastYearVS2YearsAgoWeight*Number(requestedSeatsWorksheet[amudot.TwoYRSAgoSeat+row].v);
tmp3=requestedSeatsWorksheet[amudot.TwoYRSAgoSeat+row].v;
if (isNaN(tmp3)){ 
         tmp=tmp3.split('*');
         if ( !tmp[1]){tmp1=Number(tmp[0])} else tmp1=Number(tmp[1]);}
			else tmp1=tmp;	
			if( ! delLeadingBlnks(tmp3) )tmp1=0;
MakomMinus2Yrs=lastYearVS2YearsAgoWeight*tmp1;

tmp3=requestedSeatsWorksheet[amudot.lstYrSeat+row].v;
if (isNaN(tmp3)){ 
         tmp=tmp3.split('*');
         if ( !tmp[1]){tmp1=Number(tmp[0])} else tmp1=Number(tmp[1]);}
			else tmp1=tmp;
			if( ! delLeadingBlnks(tmp3) )tmp1=0;
MakomMinus1Yrs=lastYearVS2YearsAgoWeight*lastYearVS2YearsAgoWeight*tmp1;

part3=satisfactionHistoryWeight*(MakomMinus3Yrs+MakomMinus2Yrs+MakomMinus1Yrs);

// part mispar mekomot mevukash
numOfRequestedSeats=Number(requestedSeatsWorksheet[amudot.menRosh+row].v)+Number(requestedSeatsWorksheet[amudot.menKipur+row].v)
                  +Number(requestedSeatsWorksheet[amudot.womenRosh+row].v)+Number(requestedSeatsWorksheet[amudot.womenKipur+row].v);
part4=numberOfRequestedSeatsWeight*numOfRequestedSeats;

// part mispar mekomo mevukash vs family size
part5=requestedSeatsPerFamilySizeWeight*numOfRequestedSeats;

//sum all for first sort value
calcResult[0]=part1+part2+part3-part4-part5+10000;
//console.log('part1='+part1+' part2='+part2+' part3='+part3+' part4='+part4+' part5='+part5);

// calc nashim+gvarim issue in floor sort value

part6=personalIssueWeight*(Number(requestedSeatsWorksheet[amudot.issueInFloorWmn+row].v)+Number(requestedSeatsWorksheet[amudot.issueinFloorMen+row].v));
  
// part satisfaction history
stsfctnMinus3Yrs=getStsfctnVlu(amudot.stsfctnInFlr3YRSAgoYrWmn+row)+getStsfctnVlu(amudot.stsfctnInFlr3YRSAgoYrMen+row);
stsfctnMinus2Yrs=lastYearVS2YearsAgoWeight*(getStsfctnVlu(amudot.stsfctnInFlr2YRSAgoYrWmn+row)+getStsfctnVlu(amudot.stsfctnInFlr2YRSAgoYrMen+row));
stsfctnMinus1Yrs=lastYearVS2YearsAgoWeight*lastYearVS2YearsAgoWeight*(getStsfctnVlu(amudot.stsfctnInFlrLastYrWmn+row)+getStsfctnVlu(amudot.stsfctnInFlrLastYrMen+row));
//console.log('row='+row+'   -1='+stsfctnMinus1Yrs+'   -2='+stsfctnMinus2Yrs+'   -2='+stsfctnMinus3Yrs);

/*
//old version
tmp=(requestedSeatsWorksheet[amudot.stsfctnInFlrLastYrWmn+row].v).split('*');
if ( !tmp[1]){tmp1=Number(tmp[0])} else tmp1=Number(tmp[1]);
tmp=(requestedSeatsWorksheet[amudot.stsfctnInFlrLastYrMen+row].v).split('*');
if ( !tmp[1]){tmp2=Number(tmp[0])} else tmp2=Number(tmp[1]);
stsfctnMinus1Yrs=lastYearVS2YearsAgoWeight*lastYearVS2YearsAgoWeight*(tmp1+tmp2);
 */                      
 part7= satisfactionInFloorWeight*( stsfctnMinus3Yrs+  stsfctnMinus2Yrs+ stsfctnMinus1Yrs);
 
 // part baby
 reqtmp=delLeadingBlnks(requestedSeatsWorksheet[amudot.preferedMinyanW+row].v).substr(0,1);
 SourceBirthDate=delLeadingBlnks(requestedSeatsWorksheet[amudot.preferedExplanationW+row].v);
 if (SourceBirthDate && (reqtmp == '0')){
         
       modifiedBirthDate=SourceBirthDate.substr(3,3)+' '+SourceBirthDate.substr(0,2)+', '+SourceBirthDate.substr(7,4);
			 birthDayMilisec=Date.parse(modifiedBirthDate);
       dt=new Date();
			 now=dt.getTime()
			 babyAgeInDays=Math.floor((now-birthDayMilisec)/(1000*3600*24));
   
       if (babyAgeInDays > 730 ) {part8=0}else part8=730-babyAgeInDays;
      } 
     else part8=0;
		 
		 part8=BabyWeight*part8;
 
 //sum all for first sort value
calcResult[1]=part6+part2+part3-part7-part4-part5+part8+10000;
 //	console.log('part6='+part6+' part7='+part7+' part8='+part8);	 
return calcResult;


}

//-------------------------------------------------------------------------- 
function getStsfctnVlu(cellPtr){
var tmp=[];
var tmp1;
			

tmp1=(requestedSeatsWorksheet[cellPtr].v).toString();
if( ! delLeadingBlnks(tmp1) )return 10;
 
if (tmp1.indexOf('*')== -1)return Number(tmp1);
tmp=tmp1.split('*');
if (tmp[1])return Number(tmp[1]);
return Number(tmp[0]);
}


//-------------------------------------------------------------------------- 

function sortOrderFirstParam(a,b){
    return Number(b.split('$')[1])-Number(a.split('$')[1]);
	}

//-------------------------------------------------------------------------- 
function sortOrderSecondParam(a,b){
    return Number(b.split('$')[2])-Number(a.split('$')[2]);
	}

//-------------------------------------------------------------------------- 



  function countMenAndWomenAssignedSeats(row){  
	    tmpArry=[];
			roshSeats=countSeats( delLeadingBlnks(requestedSeatsWorksheet[amudot.assignedSeatsRosh +row].v));
			kipurSeats=countSeats(delLeadingBlnks( requestedSeatsWorksheet[amudot.assignedSeatsKipur +row].v));
			tmpArry[0]=roshSeats[0];
			tmpArry[1]=roshSeats[1];
			tmpArry[2]=kipurSeats[0];
			tmpArry[3]=kipurSeats[1];
			return tmpArry;

  } 
//--------------------------------------------------------------------------  	
	
	function countSeats(seatStr){  
	  var tArry = new Array;
		tArry[0]=0;   tArry[1]=0;  if (! seatStr) return tArry;
	  stArry=seatStr.split('+');  
		for (ii=0; ii<stArry.length; ii++){
		 seatNum=Number(stArry[ii]); 
		 if(isWoman[seatNum]){tArry[1]++;}else tArry[0]++;
		    }
		return tArry;
	}	 
//--------------------------------------------------------------------------  
  
	
	
	function  closeSeats(moed,row){  
	
	var dbgCloseSeats=[];
	var dbgCloseSeatsFlag;
	 
	var alreadyAssignedTemp = new Array;
	var ptrCol,ii,strOfSeats,seatNm,tmpAssigned,ptrNN,nameForSeat;
	 
	dbgCloseSeatsFlag=false;   if( dbgCloseSeats.indexOf(row) != -1)dbgCloseSeatsFlag=true; 
	ptrNN=amudot.name+row;
	nameForSeat=requestedSeatsWorksheet[ptrNN].v;
	if( moed ==1) {
	           ptrCol=amudot.assignedSeatsRosh;
						  alreadyAssignedTemp=alreadyAssignedSeatsRosh;
							   }
	        else { 
					      ptrCol=amudot.assignedSeatsKipur;
	              alreadyAssignedTemp=alreadyAssignedSeatsKipur;
								}
	for (ii = 1; ii < 1500; ii++)if (alreadyAssignedTemp[ii] == nameForSeat){alreadyAssignedTemp[ii]=''; }// clear previous assignments
	
	strOfSeats=requestedSeatsWorksheet[ptrCol+row].v;  
	if (  delLeadingBlnks(strOfSeats)){ 
	    tmpAssigned=strOfSeats.split('+');  
	    for (ii=0; ii < tmpAssigned.length; ii++){
	       seatNm=Number(tmpAssigned[ii]);  
	       alreadyAssignedTemp[seatNm]=nameForSeat;
	   };	
	
	tArray=[];
	tArray=countMenAndWomenAssignedSeats(row);
	
	if(false)console.log('row='+row+' name='+requestedSeatsWorksheet[amudot.name+row].v+'    tArray='+tArray);   //dbgCloseSeatsFlag
	
	requestedSeatsWorksheet[amudot.numberOfAssignedSeatsRoshMen+row].v=tArray[0];
	requestedSeatsWorksheet[amudot.numberOfAssignedSeatsRoshWomen+row].v=tArray[1];
	requestedSeatsWorksheet[amudot.numberOfAssignedSeatsKipurMen+row].v=tArray[2];
	requestedSeatsWorksheet[amudot.numberOfAssignedSeatsKipurWomen+row].v=tArray[3];
	}
	update_namesForSeat(row);
	
	}
//-------------------------------------------------------------------------- 
//var assignedPerUlam=[[[0,0],[0,0]],[[0,0],[0,0]]];  //[ulam][moed][gvarim-nashim]

function CountAssignedPerMoed_PerUlam(){
 
 for (k=0;k<2;k++) //ulam
  for (kk=0;kk<2;kk++) // moed
	  for (kkk=0;kkk<2;kkk++)assignedPerUlam[k][kk][kkk]=0;   // clear counters

	
	 for (member=firstSeatRow; member<lastSeatRow+1; member++){ 
	       countAssignedPerMoed(0,member);  // rosh
 	       countAssignedPerMoed(1,member);  // kipur
      }
  

}
//--------------------------------------------------------------------------   
function countAssignedPerMoed(moed,mmbr){
  var assgndColomns=[amudot.assignedSeatsRosh,amudot.assignedSeatsKipur];
  ptr=assgndColomns[moed]+mmbr.toString(); 
	
	assgndSeatsTemp=delLeadingBlnks(requestedSeatsWorksheet[ptr].v);   
	if ( ! assgndSeatsTemp)return;
	
	
	assgndSeats=assgndSeatsTemp.split('+');
	for (j=0; j< assgndSeats.length; j++){
	   Cseat=Number(assgndSeats[j]);
		 gender=isWoman[Cseat];
	//	 eizorAndGender=getEizorForSeat(Cseat);
		 eizor=['u','m'].indexOf(UlamMartef[Cseat]);
		 assignedPerUlam[eizor][moed][gender]++;
	}
	
}



//-------------------------------------------------------------------------- 

app.get('/getAssignmentReport', function(req, res) {
	res.header("Access-Control-Allow-Origin", "*");
	 res.setHeader('Content-Type', 'text/html');
	 passW=decodeURI(req.originalUrl).split('?')[1];
	 if (passW != mngmntPASSW){
	      res.send('---');
				return;
				}
				
	rtrnStr='';
	for (i=firstSeatRow;i<lastSeatRow+1;i++){ 
      row=i.toString();
			cell=requestedSeatsWorksheet[amudot.name+row];
			if (! cell)continue;
			closeSeats(1,row); // recalculate numberOfAssignedSeats
			closeSeats(2,row);
						
			req_men_rosh=Number(requestedSeatsWorksheet[amudot.menRosh+row].v);
			req_wmn_rosh=Number(requestedSeatsWorksheet[amudot.womenRosh+row].v);
			req_men_kipur=Number(requestedSeatsWorksheet[amudot.menKipur+row].v);
			req_wmn_kipur=Number(requestedSeatsWorksheet[amudot.womenKipur+row].v);
			asgnd_men_rosh=Number(requestedSeatsWorksheet[amudot.numberOfAssignedSeatsRoshMen+row].v);
			asgnd_wmn_rosh=Number(requestedSeatsWorksheet[amudot.numberOfAssignedSeatsRoshWomen+row].v);
			asgnd_men_kipur=Number(requestedSeatsWorksheet[amudot.numberOfAssignedSeatsKipurMen+row].v);
			asgnd_wmn_kipur=Number(requestedSeatsWorksheet[amudot.numberOfAssignedSeatsKipurWomen+row].v);
					
							

      d_men_rosh=req_men_rosh-asgnd_men_rosh;
	    d_wmn_rosh=req_wmn_rosh-asgnd_wmn_rosh;
	    d_men_kipur=req_men_kipur-asgnd_men_kipur;
	    d_wmn_kipur=req_wmn_kipur-asgnd_wmn_kipur;
      nam=simplifyName(requestedSeatsWorksheet[amudot.name+row].v);
			/*nam=nam.split('*');
			if ( nam[1]  && nam[2] ){tmp=nam[1]+' '+hebrewLetters.vav+nam[2] } else tmp=nam[1]+nam[2];
			nam=nam[0]+' '+tmp;
			*/
			if( d_men_rosh || d_wmn_rosh  || d_men_kipur || d_wmn_kipur )   // at least one of them do not match
			     rtrnStr=rtrnStr+nam+'&'+req_men_rosh.toString()+'&'+ req_wmn_rosh.toString()+ '&'+req_men_kipur.toString()+'&'+req_wmn_kipur.toString()
				    +'&'+asgnd_men_rosh.toString()+'&'+asgnd_wmn_rosh.toString()+'&'+asgnd_men_kipur.toString()+'&'+asgnd_wmn_kipur.toString()+'$';
		}  // for
		if (	rtrnStr )rtrnStr=rtrnStr.substr(0,rtrnStr.length-1);	
			
			res.send('+++' + rtrnStr);
			
		})	
				         

//--------------------------------------------------------------------------    
function saveActionLog(inpStr){









}

//--------------------------------------------------------------------------    

	function sendMsgToKehilatArielSeatsGmail(titl,Msg){
	
		 var mailOptions = {
    from: 'kehilatarielseats@outlook.com', // sender address
    to: 'kehilatarielseats@gmail.com', // list of receivers
    subject: titl, // Subject line
    text: Msg //, // plaintext body
                       };
									    
   transporter.sendMail(mailOptions, function(error, info){
    if(error)  console.log('send mail (title='+titl+')  reported an error=='+error);
	    })
}
//--------------------------------------------------------------------------    
	function sendErrorToKehilatArielSeatsGmail(errMsg){
	textTosend=errMsg+'    inputString='+	 inputString;
	title='error message';
	sendMsgToKehilatArielSeatsGmail(title,textTosend);
			}
//----------------------------------------------------------------------------------

app.get('/dnldRequests', function(req, res) {
	res.header("Access-Control-Allow-Origin", "*");
	
	inputString=decodeURI(req.originalUrl); 
  passW=inputString.split('-')[1]; //tt=inputString.split('-'); 
	if (passW == mngmntPASSW){  fileToSendName= 'membersRequests.xlsx';  fileToRead=XLSXfilename}
				else {fileToSendName='empty.xlsx'; fileToRead='empty.xlsx'}
				
				
				
        res.setHeader('Content-type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
        res.setHeader('Content-Disposition', 'attachment; filename=' + fileToSendName);
	      var fileR = fs.readFileSync(fileToRead, 'binary');
        res.setHeader('Content-Length', fileR.length);
        res.write(fileR, 'binary');
        res.end();
      
 
});

//----------------------------------------------------------------------------------

app.get('/dnldSupportTblsFile', function(req, res) {
	res.header("Access-Control-Allow-Origin", "*");
	
	inputString=decodeURI(req.originalUrl); 
  passW=inputString.split('-')[1]; //tt=inputString.split('-'); 
	if (passW == mngmntPASSW){  fileToSendName= 'supportTables.xlsx';  fileToRead=supportTblsFilename}
				else {fileToSendName='empty.xlsx'; fileToRead='empty.xlsx'}
				
				
				
        res.setHeader('Content-type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
        res.setHeader('Content-Disposition', 'attachment; filename=' + fileToSendName);
	      var fileR = fs.readFileSync(fileToRead, 'binary');
        res.setHeader('Content-Length', fileR.length);
        res.write(fileR, 'binary');
        res.end();
      
 
});
//----------------------------------------------------------------------------------

app.get('/seatsOrderedXLS', function(req, res) {
	res.header("Access-Control-Allow-Origin", "*");
	
	inputString=decodeURI(req.originalUrl); 
  tmp=inputString.split('-')[1]; //tt=inputString.split('-'); 
	tmp=tmp.split('$');
	passW=tmp[0];
	DST_inIsrael=Number(tmp[1]);
	if (passW == mngmntPASSW){  fileToSendName= 'seatsOrdered.xlsx';  fileToRead=seatsOrderedFileName; generate_seatsOrderedXLS(DST_inIsrael);}
				else {fileToSendName='empty.xlsx'; fileToRead='empty.xlsx'}
				
				/*  send file to mail for debug purpose
				 var mailOptions = {
            from: 'kehilatarielseats@gmail.com', // sender address
            to: 'kehilatarielseats@gmail.com', // list of receivers
            subject: 'seatorderedlist', // Subject line
            text: 'debug',  // plaintext body
            attachments: [
                {   // file on disk as an attachment
               filename: 'seatsOrdered.xlsx',
               path: seatsOrderedFileName // stream this file
              }  ]                  
											
					};
			console.log('sending seatsOrdered.xlsx');						    
     transporter.sendMail(mailOptions, function(error, info){
         if(error)  console.log('send seatsorderedfile mail reported an error=='+error);
	    })
			*/
				
        res.setHeader('Content-type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
        res.setHeader('Content-Disposition', 'attachment; filename=' + fileToSendName);
	      var fileR = fs.readFileSync(fileToRead, 'binary');
        res.setHeader('Content-Length', fileR.length);
        res.write(fileR, 'binary');
        res.end();
      
 
});
//---------------------------------------------------------------------------------	
 //emptyHazmanatmekomotFileName=	process.env.OPENSHIFT_DATA_DIR+'hazmanatMekomotEmpty.xlsx';  
	 function generate_seatsOrderedXLS(DST_inIsrael){
	 initFromFiles('');
	 var firstRowInHazmanot=12;
	 var memberDataName =new Array; 
	 memberDataName=[];
	 var memberDataMenR =new Array; 
	 memberDataMenR=[];
	 var memberDataWmnR =new Array; 
	 memberDataWmnR=[];
	 var memberDataMenK =new Array; 
	 memberDataMenK=[];
	 var memberDataWmnK =new Array; 
	 memberDataWmnK=[];
	 
	 var nameslist = new Array;
	 var amudotHazmana=['A','B','C','D','E','F','G','H','I','J','K','L','M','N','O'];
	 idx=0;
	 for(ijk=0;ijk<familyNames.length;ijk++){
	    if( ! delLeadingBlnks( familyNames[ijk] ))continue;
			if ( hisName[ijk]  && herName[ijk]){ nameslist[idx]=familyNames[ijk]+' '+hisName[ijk]+' '+hebrewLetters.vav+herName[ijk]}
			    else nameslist[idx]=familyNames[ijk]+' '+hisName[ijk]+herName[ijk];
	   
	    idx++;
			};
	 nameslist= nameslist.sort();
	//console.log('2127 nameslist='+nameslist);
	 for (ik=0; ik<nameslist.length;ik++){
	 //  console.log('ik='+ik+' nameslist[ik='+nameslist[ik]);
	     memberDataName[ik]=nameslist[ik];
			 tmp=knownName(memberDataName[ik]); 
			 rowNum=tmp[0];
	     roww=rowNum.toString();
			 ptr=amudot.menRosh+roww;
	     memberDataMenR[ik]=requestedSeatsWorksheet[ptr].v;
			 ptr=amudot.menKipur+roww;
	     memberDataMenK[ik]=requestedSeatsWorksheet[ptr].v;
			 ptr=amudot.womenRosh+roww;
	     memberDataWmnR[ik]=requestedSeatsWorksheet[ptr].v;
			 ptr=amudot.womenKipur+roww;
	     memberDataWmnK[ik]=requestedSeatsWorksheet[ptr].v;
			
			}
			
	
			//read empty xls file
		
		tmpfile=xlsx.readFile('hazmanatMekomotEmpty.xlsx');  
		hazmanotSheet= tmpfile.Sheets['mekomot'];
		currentRow=	firstRowInHazmanot-1;
	 // fill it with info
	    third=Math.round(memberDataName.length/3+0.4);
	    for (ik=0; ik<third;ik++){
			   currentRow++;
				 roww=currentRow.toString();
			   nextColNum=0;
	        for (ikk=ik; ikk<memberDataName.length;ikk=ikk+third){
		         
					 ptr=amudotHazmana[nextColNum]+roww;
					 nextColNum++;
					 hazmanotSheet[ptr].v=memberDataName[ikk];
					
					 ptr=amudotHazmana[nextColNum]+roww;
					 nextColNum++;
					 hazmanotSheet[ptr].v=memberDataMenR[ikk];
					
					 ptr=amudotHazmana[nextColNum]+roww;
					 nextColNum++;
					 hazmanotSheet[ptr].v=memberDataWmnR[ikk];
					
					 ptr=amudotHazmana[nextColNum]+roww;
					 nextColNum++;
					 hazmanotSheet[ptr].v=memberDataMenK[ikk];
					
					 ptr=amudotHazmana[nextColNum]+roww;
					 nextColNum++;
					 hazmanotSheet[ptr].v=memberDataWmnK[ikk];
					 }  // for ikk
			} // for ik		 
		
		
		
			hazmanotSheet['B4'].v=getPrintableDate(DST_inIsrael);
		
			
	 // write the data into a new file
	 
	 xlsx.writeFile(tmpfile, seatsOrderedFileName);
	 
	 }
//---------------------------------------------------------------------------------------- 

function generate_registeredList_XLS(DST_inIsrael){
	 initFromFiles('');
	 var firstRowInRegistered=12;
	 var firstRowInNotRegistered=80;
	
	 registeredMembersName=[];
	 not_registeredMembersName=[];
	 var k=0;
	 var l=0;
	 var amudotRegistered=['A','B','C','D'];
	 var nameslist = new Array;
	 var name;
	
	 for(ijk=0;ijk<familyNames.length;ijk++)
	       if ( hisName[ijk]  && herName[ijk] ){nameslist[ijk]=familyNames[ijk]+' '+hisName[ijk]+' '+hebrewLetters.vav+herName[ijk] } else nameslist[ijk]=familyNames[ijk]+' '+hisName[ijk]+herName[ijk];
	
	 nameslist.sort();
	
	 for (ik=0; ik<nameslist.length;ik++){
	  
		 name=nameslist[ik];  
		 if ( ! delLeadingBlnks(name) )continue;
		 rowNum=knownName(name)[0];
	   roww=rowNum.toString();
	  
	   if(Number( delLeadingBlnks(requestedSeatsWorksheet[amudot.tashlum+roww].v) ) ){  // tashlum not zero meand he has registered
		    registeredMembersName[k]=name;
				k++}
			else {	
			  not_registeredMembersName[l]=name;
			  l++;
			}
	 
	    
		} // for ik
			
	
			//read empty xls file
		
		tmpfile=xlsx.readFile('registeredListEmpty.xlsx');  
		registeredSheet= tmpfile.Sheets['mekomot'];
		currentRow=	firstRowInRegistered-1;
		
	 // fill it with info
	    fourth=Math.round(registeredMembersName.length/4+0.3);
			

	    for (ik=0; ik<fourth;ik++){  // fill registered
			   currentRow++;
				 roww=currentRow.toString();
			   nextColNum=0;
	        for (ikk=ik; ikk<registeredMembersName.length;ikk=ikk+fourth){
		         
					  ptr=amudotRegistered[nextColNum]+roww;
					  nextColNum++;
					  registeredSheet[ptr].v=registeredMembersName[ikk];
					
					
					 }  // for ikk
			} // for ik		 


		// now fill not registered
		currentRow=	firstRowInNotRegistered-1;
		fourth=Math.round(not_registeredMembersName.length/4+0.3);
		for (ik=0; ik<fourth;ik++){  
			   currentRow++;
				 roww=currentRow.toString();
			   nextColNum=0;
	        for (ikk=ik; ikk<not_registeredMembersName.length;ikk=ikk+fourth){
		         
					  ptr=amudotRegistered[nextColNum]+roww;
					  nextColNum++;
					  registeredSheet[ptr].v=not_registeredMembersName[ikk];
					
					
					 }  // for ikk
			} 
			
			
		// put date and time in the xls file
		
			
			registeredSheet['B4'].v=getPrintableDate(DST_inIsrael);
			
			
	 // write the data into a new file
	 
	 xlsx.writeFile(tmpfile, registeredListFilename);
	 
	 }


//---------------------------------------------------------------------------------	 
 function getPrintableDate(DST_inIsrael){
   var offset=0;
	 var localTimeZoneDiffToZero,dParts,HR,dy,mnth,yr,newDate,newTime;
		var d= Date();
		var mnthLngth=[31,28,31,30,31,30,31,31,30,31,30,31];	
		var monthNames=['Jan','Feb','Mar','Apr','May','Jun','Jul','Aug','Sep','Oct','Nov','Dec'];	 
			var n = new Date; 	
			dParts=n.toString().split(' '); 
			localTimeZoneDiffToZero= Number(dParts[5].substr(4,4));   
			offset=2+DST_inIsrael-localTimeZoneDiffToZero; //  Israel is GMT+2
			HR=Number(dParts[4].substr(0,2))+offset;
			dy=Number(dParts[2]);
			mnth=monthNames.indexOf(dParts[1])+1;
			yr=Number(dParts[3]);
			if (yr/4 == Math.round(yr/4)){mnthLngth[1]=29} else mnthLngth[1]=28;
			
			if (HR > 23){   HR=HR-24;   dy++};
			if(dy > mnthLngth[mnth-1] ) {dy=1; mnth++};
			if (mnth>12){mnth=1; yr++};
			newDate=dy.toString()+'/'+mnth.toString()+'/'+yr.toString();
			newTime=HR.toString()+dParts[4].substr(2);
			return newDate+' '+newTime;
	}
	
//---------------------------------------------------------------------------------	  var amudotForMemberInfo	 ={name:'A',zug_gever_yisha:'F',email:'G',addr:'H',phone:'I',	memberShipStatus:'AU'};						
 

app.get('/addMember', function(req, res) {
	res.header("Access-Control-Allow-Origin", "*");

	 inputString=decodeURI(req.originalUrl);  
	 inputPairs=inputString.split('?')[1];
	 inputArray=inputPairs.split('$');
	 nam=inputArray[0];
	 if ( delLeadingBlnks(requestedSeatsWorksheet[amudot.name+(lastSeatRow+1).toString() ].v) != '$$$'){ res.send('999' ); return};  // no more space for new members
 
 //  passW=inputPairs[5];
	// if (passW == mngmntPASSW){
	 initFromFiles('');
	 
	  rslt=isThisNameKnown(nam);
		if ( ! rslt[1] ){ // continue update;


		 lastSeatRow++;
		 roww=lastSeatRow.toString();
		// console.log('lastSeatRow='+lastSeatRow+ ' inputArray='+inputArray);
		 for (j=0;j< inputArray.length;j++){
					  entry=inputArray[j];
	          entryParts=entry.split('>');
						
				    switch 	(entryParts[0]) {  
				     				 
				     case 'name': 
						          requestedSeatsWorksheet[amudot.name+row ].v=entryParts[1]; 
						          break;
							case 'email':
							        requestedSeatsWorksheet[amudot.email+row ].v=entryParts[1]; 
							        break;
											
							case 'phone':
							        requestedSeatsWorksheet[amudot.phone+row ].v=entryParts[1]; 
							        break;
											
							case 'addr':
							        requestedSeatsWorksheet[amudot.addr+row ].v=entryParts[1]; 
							        break;
											
							case 'membership':
							        requestedSeatsWorksheet[amudot.memberShipStatus+row ].v=entryParts[1]; 
							     break;
									 
						}   // switch			 																
						} // for j
		 
	
	 ptr=amudot.stsfctnInFlr2YRSAgoYrWmn+roww;    // set values for sorting so that a new member will not get high priority for un-existing past
	 requestedSeatsWorksheet[ptr].v=10;
	 ptr=amudot.stsfctnInFlr2YRSAgoYrMen+roww; 
	 requestedSeatsWorksheet[ptr].v=10;
	 ptr=amudot.TwoYRSAgoSeat+roww; 
	 requestedSeatsWorksheet[ptr].v=0;
	 ptr=amudot.stsfctnInFlr3YRSAgoYrWmn+roww; 
	 requestedSeatsWorksheet[ptr].v=10;
	 ptr=amudot.stsfctnInFlr3YRSAgoYrMen+roww; 
	 requestedSeatsWorksheet[ptr].v=10;
	 ptr=amudot.ThreeYRSAgoSeat+roww; 
	 requestedSeatsWorksheet[ptr].v=0;
	 
	 sortDB();
	   
	 xlsx.writeFile(workbook, XLSXfilename);
	 
	
	 
	
						
	 res.send('+++');
	 }
	else res.send('---' ); 
	 
});
	
//---------------------------------------------------------------------------------	   

app.get('/deleteMember', function(req, res) {
	res.header("Access-Control-Allow-Origin", "*");

	 inputString=decodeURI(req.originalUrl);  
	 nam=inputString.split('?')[1];
	 for (i=firstSeatRow;i< lastSeatRow+1;i++){
	
	  if(requestedSeatsWorksheet[amudot.name+(i).toString() ]){
		 if(delLeadingBlnks(requestedSeatsWorksheet[amudot.name+(i).toString() ].v) == nam){  //found
		       roww=(i).toString();
		       Object.keys(amudot).forEach(function(key)  {   // clear values for row
											    colmn=amudot[key];
													requestedSeatsWorksheet[ colmn+roww]={t:"s",v:''};
														
													//requestedSeatsWorksheet[ colmn+roww].v=''; 
												    
													 }) ;
						requestedSeatsWorksheet[ amudot.name+roww].v='$$$';							     
            sortDB(); 
		         res.send('+++' );
						 return;
						 };  // deleted
					}	 
		  } // end for
			 res.send('---' );   // not found
	});	 
//---------------------------------------------------------------------------------var amudotForMemberInfo	 ={name:'A',zug_gever_yisha:'F',email:'G',addr:'H',phone:'I',	memberShipStatus:'AU'};				   
app.get('/modifyMemberInfo', function(req, res) {
	res.header("Access-Control-Allow-Origin", "*");


   
	 inputString=decodeURI(req.originalUrl);  
	 inputString=inputString.split('?')[1];

	 inputArray=inputString.split('$');
	 
	 if (inputArray[0] != mngmntPASSW){
	   res.send('999 wrong  password');
		 return;
		 }
		 
	 initFromFiles('');  // last year info  
	  oldName=inputSArray[1];
		
	 for (i=firstSeatRow;i< lastSeatRow+1;i++){
	   if(delLeadingBlnks(requestedSeatsWorksheet[amudot.name+(i).toString() ].v) == oldName){  //found
		       roww=(i).toString();
					
					for (j=2;j< inputArray.length;j++){
					  entry=inputArray[j];
	          entryParts=entry.split('>');
						
				    switch 	(entryParts[0]) {  
				     				 
				     case 'name': 
						          requestedSeatsWorksheet[amudot.name+row ].v=entryParts[1]; 
						          break;
							case 'email':
							        requestedSeatsWorksheet[amudot.email+row ].v=entryParts[1]; 
							        break;
											
							case 'phone':
							        requestedSeatsWorksheet[amudot.phone+row ].v=entryParts[1]; 
							        break;
											
							case 'addr':
							        requestedSeatsWorksheet[amudot.addr+row ].v=entryParts[1]; 
							        break;
											
							case 'membership':
							        requestedSeatsWorksheet[amudot.memberShipStatus+row ].v=entryParts[1]; 
							     break;
									 
						}   // switch			 																
						} // for j
						} // if
						
					
						  
			xlsx.writeFile(workbook, XLSXfilename);  // write update								 
						 
	   res.send('+++');
		 return;
		 } //  found and modified
		 
		res.send('---');   // not found, error 
 });				
//---------------------------------------------------------------------------------
 function NextColumn(colum){
   var letters=['A','B','C','D','E','F','G','H','I','J','K','L','M','N','O','P','Q','R','S','T','U','V','X','Y','Z'];
   var i,firstLetter,seconLetter;
	 if (colum.length == 1){
	    i=letters.indexOf(colum);	
	    if (i < letters.length-1)return letters[i+1];																																																						 																																											 			
      return 'AA';
			}
  firstLetter=colum.substr(0,1);
	secondLetter=colum.substr(1,1);
	i=letters.indexOf(secondLetter);
	 if (i < letters.length-1)return firstLetter+letters[i+1];	
	 i=	i=letters.indexOf(firstLetter);
	 return letters[i+1]+'A';
	 
 }
//---------------------------------------------------------------------------------		

function sortDB(){
var i,row,j;
var compressedDB=new Array;
var compressedDBEntry=new Array;



compressedDBIdx=0;
numOfDeletedRows=0;
for (i=firstSeatRow;i<lastSeatRow+1;i++){
  nextCol=amudot.name;
	row=(i).toString();
	cell=requestedSeatsWorksheet[ amudot.name+row];
	if ( ! cell){numOfDeletedRows++;  continue;};
	if (delLeadingBlnks(cell.v) == '$$$'){numOfDeletedRows++;  continue;};
	if ( ! delLeadingBlnks(cell.v) ){numOfDeletedRows++;  continue;};
  compressedDB[compressedDBIdx]='';
	
	oneColBeyondLastCol=NextColumn(lastCol); 
	while (nextCol != oneColBeyondLastCol){     // collect values for row
	    
													cell=requestedSeatsWorksheet[ nextCol+row];
													if (cell){vlu=cell.v} else vlu='';
													compressedDB[compressedDBIdx]=compressedDB[compressedDBIdx]+'<@>'+vlu;
													nextCol=NextColumn(nextCol); 
							 } ;

      compressedDB[compressedDBIdx]=compressedDB[compressedDBIdx].substr(3);  // remove first delimiter
			
			compressedDBIdx++;
			}  // for i
	

	compressedDB.sort();
	
	
	
	for(i=0; i<compressedDB.length;i++){
	  nextCol=amudot.name;
	  row=(i+firstSeatRow).toString();  
	   compressedDBEntry=compressedDB[i].split('<@>');		
    
		
		 j=0;
	   while (nextCol != oneColBeyondLastCol){     // restore values for row
	                        requestedSeatsWorksheet[ nextCol+row]={t:"s",v:compressedDBEntry[j]}; 
													nextCol=NextColumn(nextCol); 
											    j++;
													
													
							 } ;
		} // for i
		rowIdx=compressedDB.length+firstSeatRow;   
		for (i=0; i<numOfDeletedRows;i++){
		   row=rowIdx.toString();
			 requestedSeatsWorksheet[ amudot.name+row]={t:"s",v:'$$$'};  // use empty rows for future additions
			 rowIdx++;
			 }
		
			xlsx.writeFile(workbook, XLSXfilename);  // write update 
			
			initFromFiles('');
			
}	 
	 	
//---------------------------------------------------------------------------------	   

	
 

//---------------------------------------------------------------------------------   //[ulam][moed][gvarim-nashim]  assignedPerUlam  

app.get('/getCountOfSeats', function(req, res) {
	res.header("Access-Control-Allow-Origin", "*");
	inputString=decodeURI(req.originalUrl); 
	inp=inputString.split('?')[1]; 
	
  initFromFiles(inp);
	
	rspns='';
	for (ulam=0; ulam<2;ulam++)      //// rashi-martef  /  gvarim-nashim
	      for (gender=0;gender<2;gender++) rspns=rspns+maxCountSeats[ulam][gender].toString()+'$';
				
	for (ulam=0; ulam<2;ulam++)             //[ulam][moed][gvarim-nashim] 
	   for (moed=0;moed<2;moed++)
		    for (gender=0;gender<2;gender++) rspns=rspns+assignedPerUlam[ulam][moed][gender].toString()+'$';	
	
	
	// count ulam assignment by commitee	
	var nasimMartef=0;
	var nashimRashi=0;
	var gvarimMartef=0;
	var gvarimRashi=0;
	var famsNotFullyAssgnRosh=0;
	var famsNotFullyAssgnKipur=0;
	
	var rosh_nashim_martef=0;
	var rosh_gvarim_martef=0;
	var rosh_nashim_rashi=0;
	var rosh_gvarim_rashi=0;
	var kipur_nashim_martef=0;
	var kipur_gvarim_martef=0;
	var kipur_nashim_rashi=0;
	var kipur_gvarim_rashi=0;
	
	var rosh_nasim_total=0;
	var rosh_gvarim_total=0;
	var kipur_nashim_total=0;
	var kipur_gvarim_total=0;

	for (i=firstSeatRow ; i<lastSeatRow+1   ; i++){
	  row=i.toString();
	 
	var	nashim_rosh=Number(requestedSeatsWorksheet[amudot.womenRosh+row].v);
	var	nashim_kipur=Number(requestedSeatsWorksheet[amudot.womenKipur+row].v);
	var	gvarim_rosh=Number(requestedSeatsWorksheet[amudot.menRosh+row].v);
	var	gvarim_kipur=Number(requestedSeatsWorksheet[amudot.menKipur+row].v);
				
	  if(Number(requestedSeatsWorksheet[amudot.nashimMuadaf+row].v) == 1){nashimRashi++; rosh_nashim_rashi+=nashim_rosh  ;  kipur_nashim_rashi+=nashim_kipur;};
		if(Number(requestedSeatsWorksheet[amudot.nashimMuadaf+row].v) == 2){nasimMartef++; rosh_nashim_martef+=nashim_rosh;  kipur_nashim_martef+=nashim_kipur;};
	  if(Number(requestedSeatsWorksheet[amudot.gvarimMuadaf+row].v) == 1){gvarimRashi++; rosh_gvarim_rashi+=gvarim_rosh;  kipur_gvarim_rashi+=gvarim_kipur;};
	  if(Number(requestedSeatsWorksheet[amudot.gvarimMuadaf+row].v) == 2){gvarimMartef++; rosh_gvarim_martef+=gvarim_rosh;  kipur_gvarim_martef+=gvarim_kipur;};
		
		rosh_nasim_total+=nashim_rosh;
	  rosh_gvarim_total+=gvarim_rosh;
	  kipur_nashim_total+=nashim_kipur;
	  kipur_gvarim_total+=gvarim_kipur ;
		
		
		if ( ! compareAssgndVSRqstd(row,1) )famsNotFullyAssgnRosh++;
		if ( ! compareAssgndVSRqstd(row,2) )famsNotFullyAssgnKipur++;
		
    }
			
	rspns=rspns+gvarimMartef.toString()+'$'+nasimMartef.toString()+'$'+gvarimRashi.toString()+'$'+nashimRashi.toString();
	
	rspns=rspns+'$'+famsNotFullyAssgnRosh.toString()+'$'+famsNotFullyAssgnKipur.toString();
	rspns=rspns+'$'+ rosh_nashim_martef.toString()+'$'+rosh_gvarim_martef.toString()+'$'+rosh_nashim_rashi.toString()+'$'+rosh_gvarim_rashi.toString();
	rspns=rspns+'$'+ kipur_nashim_martef.toString()+'$'+kipur_gvarim_martef.toString()+'$'+kipur_nashim_rashi.toString()+'$'+kipur_gvarim_rashi.toString();
  rspns=rspns+'$'+rosh_nasim_total.toString()+'$'+rosh_gvarim_total.toString()+'$'+kipur_nashim_total.toString()+'$'+kipur_gvarim_total.toString();
  
	 res.send(rspns);  
	
	 });

//---------------------------------------------------------------------------------

function compareAssgndVSRqstd(row,chag){  // return true if assignment fits request

var assgndStr, requestForNashim, requestForGvarim;
var i;
var list=[];
countassnd=[0,0];

  switch (chag){
			
				case 1:    // rosh 
				 assgndStr=delLeadingBlnks(requestedSeatsWorksheet[amudot.assignedSeatsRosh+row].v);
				 requestForNashim=Number(delLeadingBlnks(requestedSeatsWorksheet[amudot.womenRosh+row].v));
				 requestForGvarim=Number(delLeadingBlnks(requestedSeatsWorksheet[amudot.menRosh+row].v));

				 break;
				
				case 2:     // kipur
				assgndStr=delLeadingBlnks(requestedSeatsWorksheet[amudot.assignedSeatsKipur+row].v);
				requestForNashim=Number(delLeadingBlnks(requestedSeatsWorksheet[amudot.womenKipur+row].v));
				requestForGvarim=Number(delLeadingBlnks(requestedSeatsWorksheet[amudot.menKipur+row].v));
				 break;
				
				}
		if ( ! (	requestForNashim + 	requestForGvarim)  )return true;  // no request == full assignment
		if( ! assgndStr) return false;  // not assigned yet at all
		list=assgndStr.split('+');
		for (i=0; i<list.length;i++)countassnd[isWoman[Number(list[i])]]++;
		if (  (countassnd[0] < requestForGvarim) || (countassnd[1] < requestForNashim) )return false;
		
		return true;
		
}		
		

//--------------------------------------------------------------------------------- 

// Hard initialize membersRequests.xlsx file

app.get('/shira1807', function(req, res) {
	res.header("Access-Control-Allow-Origin", "*");
	if (initDone){
	shiras_backup=true;   
	backupRequests();  // backup before every thing
	};
	
	
	tmpfile=fs.readFileSync('membersRequests.xlsx');
	fs.writeFileSync(EmptyXLSXfilename, tmpfile);
	fs.writeFileSync(XLSXfilename, tmpfile);
	workbook = xlsx.readFile(XLSXfilename);
	requestedSeatsWorksheet = workbook.Sheets['HTMLRequests'];
	
	for(i=0;i<1500;i++){
	alreadyAssignedSeatsRosh[i]=' '; 
  alreadyAssignedSeatsKipur[i]=' ';
	 };
	 lastYearInit='-';
	 initFromFiles('');
	
 console.log('membersRequests file was HARD initialized');


for(i=1; i<lastSeatNumber+1; i++)seatOcuupationLevel[i]=0;    // clear and set array size 

	
	
	
	 res.setHeader('Content-Type', 'text/html'); 
	res.send(cache_get('initialize') );
	
	 
	 });

//---------------------------------------------------------------------------------	 
// send tashlum info to gizbar

 
  app.get('/tashlumim', function(req, res) {
	res.header("Access-Control-Allow-Origin", "*");
	
	
	inputString=decodeURI(req.originalUrl); 
	
	if (inputString.substr(12)==gizbarPASSW){
	initFromFiles(''); 
	listOfPayments='';
		 

	for(i=firstSeatRow;i<lastSeatRow+1;i++){
	    
	    pointerCell=amudot.name+(i).toString(); 
		 cell=requestedSeatsWorksheet[pointerCell]; 
	   if(! cell) continue;
		 nameInCell=delLeadingBlnks(cell.v);
		 if ( ! nameInCell)continue;
		// if(nameInCell.substr(nameInCell.length-1)=='*')nameInCell=nameInCell.substr(0,nameInCell.length-1);	
			if ( hisName[i-firstSeatRow]  && herName[i-firstSeatRow] ){bothNames=hisName[i-firstSeatRow]+' '+
			         hebrewLetters.vav+herName[i-firstSeatRow]; } else bothNames=hisName[i-firstSeatRow]+herName[i-firstSeatRow];								
		 listOfPayments=listOfPayments+'$'+familyNames[i-firstSeatRow]+' '+bothNames;
		  pointerCell=amudot.tashlum+(i).toString(); 
		 cell=requestedSeatsWorksheet[pointerCell];
     listOfPayments=listOfPayments+'+'+cell.v;	
		  pointerCell=amudot.tashlumPaid+(i).toString(); 
		 cell=requestedSeatsWorksheet[pointerCell];
     listOfPayments=listOfPayments+'+'+cell.v;	  
	 }
	 listOfPayments=listOfPayments.substr(1);
	
	 res.send('+++'+listOfPayments);
	}
	else  res.send('---');
	 });
//---------------------------------------------------------------------------------	 

app.get('/UPDtashlumim', function(req, res) {
	res.header("Access-Control-Allow-Origin", "*");
	
	inputString=decodeURI(req.originalUrl);  
	
	
	inputPairs=inputString.split('$'); 
	if (inputPairs[1] != gizbarPASSW){ res.send('999')}
	
	else{
	initFromFiles('');
		 errFound=false;
	 // check request validity
	 for (i=2;i<inputPairs.length;i++){
	   memberUpd=inputPairs[i].split('+');
		 paid=delLeadingBlnks(memberUpd[1]);
		 if( ! paid) continue;
		 row=knownName(memberUpd[0]);
		 if (row==-1) {errFound=true; break};
	 }	
	 if ( errFound ){ res.send('---')}
	 else{  //go update
	  	 atLeastOneToUpdate=false;
       for (i=2;i<inputPairs.length;i++){
	     memberUpd=inputPairs[i].split('+');
		   paid=delLeadingBlnks(memberUpd[1]);
		//   if( ! paid) continue;
		   row=knownName(memberUpd[0])[0]; 
	 		 ptr=amudot.tashlumPaid+row.toString();   
			 requestedSeatsWorksheet[ptr].v=paid;
			 atLeastOneToUpdate=true;
			
			 }
			 if (atLeastOneToUpdate) {xlsx.writeFile(workbook, XLSXfilename)};
			 res.send('+++');
			 }  //2nd else
	 } //first else
	 });


//---------------------------------------------------------------------------------   		
	 


// get request to write member's input
   var inputArray = new Array;
  app.get('/writeinfo', function(req, res) {
	var dbg,i;
	res.header("Access-Control-Allow-Origin", "*");
	fullInpString=decodeURI(req.originalUrl); 
	inputString=fullInpString.split('?')[1]; 
	tmp=inputString.split('>'); 
	DST_inIsrael=Number(tmp[0]);
	inputString=tmp[1];  
	dbg=searchDebugParam('writeinfo');
	if (  dbg != -1) console.log('write_info at '+getPrintableDate(DST_inIsrael)+ ' of inputString='+inputString);
	initFromFiles('');
	
	// save transaction history
	for (i=0;i<300;i++){
	  transactionPtr='A'+(transactionHistoryFirstRow+i).toString();
		if (delLeadingBlnks(requestedSeatsWorksheet[transactionPtr].v)  == '$$$'){ //empty slot found
		       requestedSeatsWorksheet[transactionPtr].v='write_info at '+getPrintableDate(DST_inIsrael)+ ' of inputString====='+inputString;
					 break;
				} // if
		} // for	
		if(i >= 300)console.log('space not found for transaction history');		 
	xlsx.writeFile(workbook, XLSXfilename);
	
//	inputString=inputString.substr(12); 
	inputPairs=inputString.split('&'); 
	namTitl=inputPairs[0].split('=')[1];
	managementRequest=false;
	sendMsgToKehilatArielSeatsGmail(namTitl,fullInpString);   
	 if(handleInput(inputPairs)){res.send('+++'+incompatibilty+'$'+hasStillToPay.toString()+afterClosingDate)}else { res.send('---'+errorNumber);};
	 
	 });
//---------------------------------------------------------------------------------
app.get('/getNextTransaction', function(req, res) {
 var msg,i,idx,row;
	res.header("Access-Control-Allow-Origin", "*");
	fullInpString=decodeURI(req.originalUrl);
	inputString=fullInpString.split('?')[1];
	inputArr=inputString.split('$');
	if  ( inputArr[0] != debugPASSW){ res.send('???'); return};
	idx=Number( inputArr[1]);
	if(i>299){res.send('---');  return}; // idx to big
	row=(transactionHistoryFirstRow+idx).toString();
	ptr='A'+row;
	msg=requestedSeatsWorksheet[ptr].v;
  res.send(msg);
	 return;
});

//---------------------------------------------------------------------------------
app.get('/isThereSuchAName', function(req, res) {
  var rNmA = new Array();  
	res.header("Access-Control-Allow-Origin", "*");
	inputString=decodeURI(req.originalUrl);
	inputPairs=inputString.split('$'); 
	if (inputPairs[1] != mngmntPASSW){ res.send('999')}
	
	else { 
	initFromFiles(inputPairs[3]);
	   srchArg=inputPairs[2];
	   srchArgLength=srchArg.length;         
	
	   nameParts=srchArg.split(' '); 
	   rslt=''; 
	
	  for (rn=0; rn<familyNames.length; rn++){
	   
	       fnm=familyNames[rn].split(' ');
			   if (fnm.length>1)familyNames[rn]=fnm.join(' '); 
			   leftPartOfNameString=familyNames[rn].substr(0,srchArgLength);
	       if(srchArg == leftPartOfNameString)  rslt=rslt+'$'+familyNames[rn];
			  }  // for 	
			if (rslt.length)rslt=rslt.substr(1); // remove first $													 
			
		  res.send(rslt);              
		 
			
     }   // else
			   
	});
//---------------------------------------------------------------------------------
	
// get request to verify family name and respond with previous inputs	

	app.get('/famname', function(req, res) {
	res.header("Access-Control-Allow-Origin", "*");
	inp=decodeURI(req.originalUrl).split('?')[1];
	initFromFiles('');
	respns=memberInfo('member',inp);
	res.send(respns);
	});
//----------------------------------------------------------------------

// get request to verify family name for mgmt

	app.get('/mngfmname', function(req, res) {
	res.header("Access-Control-Allow-Origin", "*");
	inp=decodeURI(req.originalUrl).split('?')[1];
	inpData=inp.split('&');   
	if(inpData[1] == mngmntPASSW){
	  initFromFiles(inpData[2]); 
	 respns=memberInfo('manage',inpData[0]);
	 res.send(respns);
	}
	else {res.send('999' );
	      }
	});
	
//----------------------------------------------------------
 // get request to close or open registration

	app.get('/closeOpen', function(req, res) {
	res.header("Access-Control-Allow-Origin", "*");
	inp=decodeURI(req.originalUrl).split('?')[1];
	inpData=inp.split('-');
	var d= Date();  
	ptr=amudot.registrationClosedDateNTime+'2';    
	if(inpData[1] == mngmntPASSW){
	   initFromFiles('');
	 if (inpData[0] == "close") {
	     //     dateTimeNow=Date.now();
	          requestedSeatsWorksheet[ptr].v=d;  
	                        }
				else if (inpData[0] == "open"){ 
				     requestedSeatsWorksheet[ptr].v=' ';
						 for(member=firstSeatRow; member<lastSeatRow+1; member++){
						                  ptr1=amudot.requestDate+(member).toString();
															requestedSeatsWorksheet[ptr1].v=' ';  
															           }         
				                      };
		 xlsx.writeFile(workbook, XLSXfilename);																																	
	 res.send('+++');
	}
	else {res.send('999' );
	      }
	});


//------------------------------------------------------------------------------
		
	
	app.get('/manage', function(req, res) {
	//res.header("Access-Control-Allow-Origin", "*");
	 res.setHeader('Content-Type', 'text/html');
	 
	 
	 var dbg;
		
		dbg=searchDebugParam('disable');  
		if ( ( dbg != -1 ) || ( ! initDone) ){
		          res.send(cache_get('real_index'));
							return;
							}
            res.send(cache_get('mgmt') );
        })
				
//----------------------------------------------------------   
							
app.get('/getFullList', function(req, res) {
	res.header("Access-Control-Allow-Origin", "*");
	 res.setHeader('Content-Type', 'text/html');
   
	 var name,i, ijk,ijl,inp,inpData,tmplist,listType,tmp,tmp1,ptr1;
	 var indxToMinNames=[];
	 inp=decodeURI(req.originalUrl).split('?')[1];
	inpData=inp.split('$');
	//for(i=0;i<inpData.length;i++) console.log('inpData['+i+']='+inpData[i]+'/'  );
	
	if(inpData[1] == mngmntPASSW){  
	initFromFiles(inpData[2]);
	numberOfCalculatedStsfction=0;
	
	 listType=inpData[4]; 
	tmplist=[];
	ijl=0;
	
	for(ijk=0;ijk<familyNames.length;ijk++){
	   if ( ! delLeadingBlnks(familyNames[ijk]))continue;
	   combinedName=familyNames[ijk]+'*'+hisName[ijk]+'*'+herName[ijk];
		 name = simplifyName(combinedName);      /* .split('*');	
		 if ( name[1]  && name[2] ){tmp=name[1]+' '+hebrewLetters.vav+name[2] } else tmp=name[1]+name[2];
		 name=name[0]+' '+tmp;
		 */
		 row=knownName(name)[0]; 
		 row=row.toString();
	   ptr1=amudot.ThisYRSSeat+row;
		 if ( typeof requestedSeatsWorksheet[ptr1] != 'undefined' )  { if ( delLeadingBlnks(requestedSeatsWorksheet[ptr1].v)){numberOfCalculatedStsfction++;} else  if (listType != 'pp') continue;};
	
	  	
		 if (listType=='problems'){
			  
				 
					ptr1=amudot.stsfctnInFlrThisYrWmn+row;
			
				if ( typeof requestedSeatsWorksheet[ptr1] !='undefined'){
				     tmp=delLeadingBlnks(requestedSeatsWorksheet[ptr1].v);  
				     if ( ! tmp){wmnCalculatedStsf='10' } else wmnCalculatedStsf=tmp.split('*')[0];
					} else 	wmnCalculatedStsf='10';
				
				ptr1=amudot.stsfctnInFlrThisYrMen+row;
				
				if (	typeof requestedSeatsWorksheet[ptr1] != 'undefined' ){ 
				    tmp=delLeadingBlnks(requestedSeatsWorksheet[ptr1].v);
			      if ( ! tmp){menCalculatedStsf='10' } else menCalculatedStsf=tmp.split('*')[0];
				  }else 	menCalculatedStsf='10';	
			
		    if (  (wmnCalculatedStsf=='10') && (menCalculatedStsf=='10') )continue
				
				
		  } // if listType=problems 
	  
		
		 
			
	    tmplist[ijl]=combinedName;
			ijl++;

   };
	 
	 //console.log('listType='+listType+' numberOfCalculatedStsfction='+numberOfCalculatedStsfction+' ijl='+ijl);  
   if ((listType != 'pp') && ( ! numberOfCalculatedStsfction)){res.send('---');return;} // this year satisfection values not avaiable yet
	 if ((listType != 'pp') && ( ! ijl)){res.send('---');return;} // empty list
	 tmplist.sort();
	 
	 for(i=0;i<tmplist.length;i++)indxToMinNames[i]=(knownName(simplifyName(tmplist[i]))[0]-firstSeatRow).toString();
			

	 listOfnames='+++'+	tmplist.join('$')+'@'+indxToMinNames.join('$');																													
	 res.send(listOfnames);
	}
	else { res.send('999' );
	      }
	 
	 
        })


//----------------------------------------------------------

  app.get('/getRequstorsList', function(req, res) {
	res.header("Access-Control-Allow-Origin", "*");
	 res.setHeader('Content-Type', 'text/html');
   
	 inpA=decodeURI(req.originalUrl).split('?')[1];
	// inpData=inp.split('$');
	inp=inpA.split('$'); 
	var tmpRqList=[]; 
	var indxToMinNames=[];
	var   idx=0;
	if(inp[0] == mngmntPASSW){  
	initFromFiles(inp[1]);
	    for (i=firstSeatRow;i<lastSeatRow+1;i++){
			 
              row=i.toString();
							nameCell=requestedSeatsWorksheet[amudot.name+row];
							if (! nameCell) continue;
							nam=delLeadingBlnks(nameCell.v);
							if ( ! nam)continue;
							tmpVl=Number(requestedSeatsWorksheet[amudot.menRosh+row].v)+Number(requestedSeatsWorksheet[amudot.womenRosh+row].v)
		             +Number(requestedSeatsWorksheet[amudot.menKipur+row].v)+Number(requestedSeatsWorksheet[amudot.womenKipur+row].v);
		       if ( !	tmpVl ) continue;  // no request made		
			
				   tmpRqList[idx]=	nam; 
						 
					idx++;
							
			}				
	 
	    tmpRqList.sort();
			for(i=0;i<tmpRqList.length;i++)indxToMinNames[i]=(knownName(simplifyName(tmpRqList[i]))[0]-firstSeatRow).toString();
			tmpRqListStr='+++'+tmpRqList.join('$')+'@'+indxToMinNames.join('$');
			if( ! idx )tmpRqListStr='000';   // empty list
	    res.send(tmpRqListStr);
	
	
	}
	else res.send('999' );
	      
	 	 
        })
	 

//----------------------------------------------------------


app.get('/getlist', function(req, res) {
	res.header("Access-Control-Allow-Origin", "*");
	 res.setHeader('Content-Type', 'text/html');
   
	 inp=decodeURI(req.originalUrl).split('?')[1];
	inpData=inp.split('$');
	
	if(inpData[1] == mngmntPASSW){
	// initFromFiles('');
	 moedCode=inpData[2];
	 SidurUlam=inpData[3];
	 shlavBFilter=inpData[4];
	 notPaidFilter=inpData[5];
	 doneWithFilter=inpData[6];
	 initFromFiles(inpData[7]);  
	 listOfnames='+++'+		filterAndSort();																														
	 res.send(listOfnames);
	}
	else res.send('999' );
	      
	 
	 
        })
				
//----------------------------------------------------------


app.get('/getMinLengthNames', function(req, res) {
	res.header("Access-Control-Allow-Origin", "*");
	 res.setHeader('Content-Type', 'text/html');
   
	 inpData=decodeURI(req.originalUrl).split('?')[1];
  inpData=inpData.split('$');
	initFromFiles(inpData[1]);
	
	if ( (inpData[0] == mngmntPASSW) || (inpData[0]== moshavimPASSW)  ){
	
	 if( ! minimumName.length){ res.send('---' );  return;};
	 listOfnames= '+++'+		minimumName.join('$');
	
	 																												
	 res.send(listOfnames);
	}
	else res.send('999' );
	      
	 
	 
        })
								
				
//---------------------------------------------------------- 
 app.get('/genCodeSendEmail', function(req, res) {
	res.header("Access-Control-Allow-Origin", "*");
	 res.setHeader('Content-Type', 'text/html');
   initFromFiles('');
	 mmbr=decodeURI(req.originalUrl).split('?')[1];
	 	 
	 row=knownName(mmbr)[0];
	 if (row == -1){
	     res.send('---1' );}
		else {
		   	emal=delLeadingBlnks(requestedSeatsWorksheet[amudot.email+row.toString()].v);
	      if(! emal) {res.send('---2' );}
				   else{
					     codeToSend=Math.floor(Math.random()*100000);
							 codeToSend=codeToSend.toString();
							 txtToSend=' the follwing passcode is valid for the next 20 minutes : ' + codeToSend;
               subjc='passcode to change email addr' ; 
	             sendMail(emal,subjc,txtToSend);
							 forgetList[row]=codeToSend+'$3';  // forget after 3 timer cycles
							 d= new Date;
							 console.log(d+' email password '+codeToSend+' was sent to '+emal+ ' for member='+mmbr+' rownum='+row);
							 res.send('+++' );
							 }
						}	 
	  })
//----------------------------------------------------------
 app.get('/ckEmailCode', function(req, res) {
	res.header("Access-Control-Allow-Origin", "*");
	 res.setHeader('Content-Type', 'text/html');
	 
	 inp=decodeURI(req.originalUrl).split('?')[1];
	 initFromFiles('');
	 inpData=inp.split('$');
	 mmbr=inpData[0];
	 codeToVerify=inpData[1];
	 row=knownName(mmbr)[0];
	 if (row == -1){ res.send('---3' );}
	   else {
		    if ( ! forgetList[row] ) {res.send('---4' );}
					  else {
						  if (  forgetList[row].split('$')[0] == codeToVerify) { res.send('+++' );} else  res.send('---5' );
				         }
					}			 
	
	   })	 
//----------------------------------------------------------
 app.get('/savePrsnlPrblmsParams', function(req, res) {
	res.header("Access-Control-Allow-Origin", "*");
	 res.setHeader('Content-Type', 'text/html');
   var tmp;
	 inp=decodeURI(req.originalUrl).split('?')[1];
	 inpData=inp.split('$');
	 
	if(inpData[0] == mngmntPASSW){
	   initFromFiles('');
	   rowNum= knownName(inpData[1])[0];   
		 if(rowNum == -1 ){res.send('---')}
		 else {
		 row=rowNum.toString();
		 
		 ptr= amudot.issueBetweenFloors + row;
		 requestedSeatsWorksheet[ptr].v=inpData[2];   // issueBetweenFloors  
		
     ptr= amudot.issueinFloorMen + row;
		 requestedSeatsWorksheet[ptr].v=inpData[3];   // issueinFloorMen
		
		 ptr= amudot.issueInFloorWmn + row;
		 requestedSeatsWorksheet[ptr].v=inpData[4];   // issueInFloorWmn
		
		 ptr= amudot.nashimMuadaf + row;
		 requestedSeatsWorksheet[ptr].v=inpData[5]; 
		 
		 ptr= amudot.gvarimMuadaf + row;
		 requestedSeatsWorksheet[ptr].v=inpData[6]; 
		 
		
		  xlsx.writeFile(workbook, XLSXfilename);																													
	   res.send('+++');
	   }
	}
	else res.send('999' );
	 
        })

//----------------------------------------------------------
 app.get('/saveStsfctionParams', function(req, res) {
	res.header("Access-Control-Allow-Origin", "*");
	 res.setHeader('Content-Type', 'text/html');
   var tmp;
	 inp=decodeURI(req.originalUrl).split('?')[1];
	 inpData=inp.split('$');
	 
	if(inpData[0] == mngmntPASSW){
	   initFromFiles('');
	   rowNum= knownName(inpData[1])[0];   
		 if(rowNum == -1 ){res.send('---')}
		 else {
		 row=rowNum.toString();
		 
		 
		 ptr= amudot.stsfctnInFlrLastYrMen + row; 
		 tmp= (requestedSeatsWorksheet[ptr].v).split('*');
		 requestedSeatsWorksheet[ptr].v=tmp[0]+'*'+inpData[2];   //lastYearStsfctnMen
		
		 ptr= amudot.stsfctnInFlrLastYrWmn + row; 
		 tmp= (requestedSeatsWorksheet[ptr].v).split('*');  
		 requestedSeatsWorksheet[ptr].v=tmp[0]+'*'+inpData[3];   //lastYearStsfctnWmn
		 
		 ptr= amudot.lstYrSeat + row;
		 tmp= (requestedSeatsWorksheet[ptr].v).split('*');
		 requestedSeatsWorksheet[ptr].v=tmp[0]+'*'+inpData[4]; //lastYearSeat
		 
		
		
		  xlsx.writeFile(workbook, XLSXfilename);																													
	   res.send('+++');
	   }
	}
	else res.send('999' );
	 
        })
//----------------------------------------------------------
						
 app.get('/saveStsfctionParams_thisYear', function(req, res) {
	res.header("Access-Control-Allow-Origin", "*");
	 res.setHeader('Content-Type', 'text/html');
   var tmp;
	 inp=decodeURI(req.originalUrl).split('?')[1];
	 inpData=inp.split('$');
	 
	if(inpData[0] == mngmntPASSW){
	   initFromFiles('');
	   rowNum= knownName(inpData[1])[0];   
		 if(rowNum == -1 ){res.send('---')}
		 else {
		 row=rowNum.toString();
		 
		 
		 ptr= amudot.stsfctnInFlrThisYrMen + row; 
		 tmp= (requestedSeatsWorksheet[ptr].v).split('*');
		 requestedSeatsWorksheet[ptr].v=tmp[0]+'*'+inpData[2];   //lastYearStsfctnMen
		
		 ptr= amudot.stsfctnInFlrThisYrWmn + row; 
		 tmp= (requestedSeatsWorksheet[ptr].v).split('*');  
		 requestedSeatsWorksheet[ptr].v=tmp[0]+'*'+inpData[3];   //lastYearStsfctnWmn
		 
		 ptr= amudot.ThisYRSSeat + row;
		 tmp= (requestedSeatsWorksheet[ptr].v).split('*');
		 requestedSeatsWorksheet[ptr].v=tmp[0]+'*'+inpData[4]; //lastYearSeat
		 
		
		
		  xlsx.writeFile(workbook, XLSXfilename);																													
	   res.send('+++');
	   }
	}
	else res.send('999' );
	 
        })
//---------------------------------------------------------- 
					
function analyseRqstVSAssgnd(rowp){
var dbg1=false;
var dbgIdx, tmp
debugRows=[];
dbgIdx=searchDebugParam('stsfctn');
if(dbgIdx != -1){
   tmp=debugRequestsAll[dbgIdx];
	 tmp=tmp.split('%')[1];  // leave only real params
	 tmp=tmp.split('$');   
	 debugRows=tmp[3].split('+');  // [3] is list of rows seperated by '+'
	 }

  	dbgStsfction=false; if(debugRows.indexOf(rowp) != -1 )dbgStsfction=true;
		
		if( [].indexOf(rowp) != -1)dbg1=true;
  var tmp,tmp0, tmp1, tmp2, tmp3;		
	var i;
	var row;
	var seatVlus=[];
	var stsArry=[];
	var tmpResults=[];
	row=rowp.toString();
			// in this process gender=0 is men =1 is women	 
		 
	//console.log('===================   analyse row='+row+'  =========');
		 	
	 
	 
	 
	 // init values
	
	
	//-------------------
	var 	wmn_mrtf_toIndex=[ [2,0],[3,1]];
	aChagRslts=[10,10,0];
  

  var originalRequestArray=(delLeadingBlnks(requestedSeatsWorksheet[amudot.markedSeats+row].v)).split('+');
	for (i=0; i<originalRequestArray.length;i++){
	  seatVlus=originalRequestArray[i].split('_');
		originalReqSeats[i]=seatVlus[0];
		originalReqSeatPriority[i]=seatVlus[1];
	}	
	
	var martefUlamGrade;
	var nameForRow, tmpPosition;
	var numOfGenders, firtGender,lastGender;
	tmp=0;
	chagimCounter=2;
	 assignedSeatslist=[[[],[]],[[],[]]];
   assgndRows=[ [[],[]], [[],[]] ];
   chagimWithRequest=[0,1];
for(chag=0;chag<2;chag++){

	 martefUlam_gvrimNashim=[[],[],[],[]];  //martef gvrim, mrtef nashim, ulam g ulam n
	 
	 counters_martefUlam_gvrimNashim=[0,0,0,0];
	 
	 ptr=amudot.assignedSeatsRosh+row;  if (chag==1)ptr=amudot.assignedSeatsKipur+row;
			 
   sts=delLeadingBlnks(requestedSeatsWorksheet[ptr].v);  
   if (! sts){     // no assignment no complaint   %%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
	       chagimCounter--;
				 tmpPosition=chagimWithRequest.indexOf(chag);
				 chagimWithRequest.splice(tmpPosition,1);  // no request for this chag
				 continue;
				 }  
 
  // get list of requested seats, full list and per ulam/gender lists
	
	 stsArry=sts.split('+');
	 
	 for(iill=0;iill<stsArry.length;iill++){
	   st=stsArry[iill];
     stN=Number(st);  
     isw=isWoman[stN];  // isw==1 ===> nashim
		 isInMrtf=0;    if(UlamMartef[stN]=='m')isInMrtf=1; 
		 indx=wmn_mrtf_toIndex[isw][isInMrtf];  
     assignedSeatslist[isw][chag].push(stN);	
		martefUlam_gvrimNashim[indx].push(st);
		counters_martefUlam_gvrimNashim[indx]++;
		      
		  roww=configRowOfSeat[stN];
			if ( assgndRows[isw][chag].indexOf(roww) == -1) assgndRows[isw][chag].push(roww); // isw==1 ===> nashim
			
	};	 // for

// calculate ulam/martef grade		per holiday	
			genderCounter=2;
		index1=wmn_mrtf_toIndex[1];
		divideBy=counters_martefUlam_gvrimNashim[index1[1]]+ counters_martefUlam_gvrimNashim[index1[0]];
		if (divideBy==0){gradeNashim=0;genderCounter--;} // no request no complaint
		  else  gradeNashim=counters_martefUlam_gvrimNashim[index1[1]]/divideBy;
    
		index1=wmn_mrtf_toIndex[0];
	  divideBy=counters_martefUlam_gvrimNashim[index1[1]]+ counters_martefUlam_gvrimNashim[index1[0]];
		if (divideBy==0){gradeGvarim=0;genderCounter--;} // no assignment no complaint
		  else  gradeGvarim=counters_martefUlam_gvrimNashim[index1[1]]/divideBy;
			
	
		if ( ! genderCounter){   //no assignment for this chag
		       chagimCounter--;  // this chag does not count for grade;
					 chagimWithRequest=chagimWithRequest.splice(chagimWithRequest.indexOf(chag),1);  // no request for this chag
					 combinedGrade=0;
					 }   else  combinedGrade= (gradeNashim+gradeGvarim)/genderCounter;
		tmp=tmp+combinedGrade;
	}	


   if(chagimCounter){martefUlamGrade=tmp/chagimCounter} else martefUlamGrade=0;
	 
	martefUlamGrade=martefUlamGrade.toString();
			
nameForRow=delLeadingBlnks(requestedSeatsWorksheet[amudot.name+row].v);
	if (nameForRow.substr(nameForRow.length-1,1)=='*')nameForRow=nameForRow.substr(0,nameForRow.length-1);	
	// get lists of requested seats for women and for men 
	
 sts=delLeadingBlnks(requestedSeatsWorksheet[amudot.markedSeats+row].v);  
 
 if (! sts) { stsfctionColction.push(row+'$10$10$'+martefUlamGrade);  return};  // no request no complaint  
  stsArry=sts.split('+');
	
	
  rqstdSeats=[[],[]];
	var prvsTmp=[[],[]];
	
 for (priority=1;priority<4;priority++){  
   
	  rqstdSeats_tmp=[[],[]];  
	rqstdRows=[[[],[]],   [[],[]]];  // nashim [row,length} gvarim [row,length]
	 
	 var numSeats=[];    
   for(iill=0;iill<stsArry.length;iill++){
	   seatVlus=stsArry[iill].split('_');
	   if (Number(seatVlus[1]) != priority )continue;
	   st=seatVlus[0];
     stN=Number(st);
		 isw= isWoman[stN] ;         // isw==1 ===> nashim
		  rqstdSeats_tmp[isw].push(stN);  
	};   // for iill
			
			 for (gender=0;gender<2;gender++) numSeats[gender]=rqstdSeats_tmp[gender].length;

			 for (gender=0; gender<2; gender++){
			   if( ( ! numSeats[gender]) || isExpansion(prvsTmp,rqstdSeats_tmp,gender,dbg1)  ){
				if(dbgStsfction) console.log('expnsn  row='+row);
			     concatArrays(rqstdSeats[gender],rqstdSeats_tmp[gender]);
										 
					 } else {
					   rqstdSeats[gender]=[];
						if(dbgStsfction)console.log('not expnsion row='+row);
						   for(iill=0;iill<rqstdSeats_tmp[gender].length;iill++)rqstdSeats[gender][iill]=rqstdSeats_tmp[gender][iill];
					} // else
					
				}  //gender	
				
				
		// save this temp for next expansion checking
	
		prvsTmp=[[],[]];
		for (gender=0;gender<2;gender++){
				for(iill=0;iill<rqstdSeats_tmp[gender].length;iill++)prvsTmp[gender][iill]=rqstdSeats_tmp[gender][iill];					 
			    }
	for (gender=0;gender<2;gender++){		
	 for(iill=0;iill<rqstdSeats[gender].length;iill++){
	    stN=rqstdSeats[gender][iill];
			roww=configRowOfSeat[stN];
			tmp0=rqstdRows[gender][0].indexOf(roww);
			if ( tmp0 == -1){
			     rqstdRows[gender][0].push(roww);
					 tmp0=rqstdRows[gender][0].indexOf(roww);
					 rqstdRows[gender][1][tmp0] =1;
					 } else rqstdRows[gender][1][tmp0]++;
					
				  
		};   // for iill
		} // for gender
		
						 
						 sortRqstdRows(0);
						 sortRqstdRows[1];
		//	 };
	
		tmp0=0;  tmp1=0;  tmp2=0; 
		
	for(chag=0;chag<2;chag++){
	  if (chagimWithRequest.indexOf(chag) != -1){  
		
	        evalRqstVSAssgnd(chag,row);   
	
					tmp0=tmp0+aChagRslts[0];  
					tmp1=tmp1+aChagRslts[1];
					}
													
	}  //for
		if (chagimCounter){
	tmp0=((tmp0/chagimCounter).toString()).substr(0,4);; 
	tmp1=((tmp1/chagimCounter).toString()).substr(0,4);
	} ; 
		
	  res1=tmp0+'$'+tmp1; 
	tmpResults[priority]=res1;  
	if(dbgStsfction) console.log('                               priority='+priority+ '  tmpResults[priority]='+tmpResults[priority]);
	}  // priority
	
	
	// choose best results
	res1=nameForRow;
	tmp0=0;
	for (priority=1;priority<4;priority++){  // choose best 1st grade 
	  tmp1=Number(tmpResults[priority].split('$')[0]);
		if (tmp1>tmp0)tmp0=tmp1;
		}
		res1=res1+'$'+tmp0.toString();
		
		tmp0=0;
	for (priority=1;priority<4;priority++){  // choose best 2nd grade 
	  tmp1=Number(tmpResults[priority].split('$')[1]);
		if (tmp1>tmp0)tmp0=tmp1;
		}
		res1=res1+'$'+tmp0.toString();
		
		res1=res1+'$'+martefUlamGrade+'$'+row;  // row is saved for debug
	
	
	stsfctionColction.push(res1);
	
	
	}
	
//----------------------------------------------------------	
	
	function sortRqstdRows(gender){
	  var temp=[];
		var tmp1=[];
		var i;
		
		for (i=0; i<rqstdRows[gender][0].length; i++)temp[i]=rqstdRows[gender][0][i].toString()+'$'+rqstdRows[gender][1][i].toString();
		temp.sort(rqstdRowsSort);
		for (i=0; i<rqstdRows[gender][0].length; i++){
		              temp1=temp[i].split('$');
		              rqstdRows[gender][0][i]=Number ( temp1[0]);
									rqstdRows[gender][1][i]=Number ( temp1[1]);
									
									}
			
							
			}						

//----------------------------------------------------------	
function rqstdRowsSort(a,b){
var t1,t2;
t1=Number(a.split('$')[1]);
t2=Number(b.split('$')[1]);
return t2-t1;
}
 
//----------------------------------------------------------	
function isExpansion(prvsTmp,rqstdSeats_tmp,gender,dbg){
var sameRow=false;

var i,j,tmp;
//if(dbg)console.log('gender='+gender+' prvsTmp[gender='+prvsTmp[gender]+' rqstdSeats_tmp[gender]='+rqstdSeats_tmp[gender]);   
  
	for (i=0; i<prvsTmp[gender].length;i++){
	  for(j=0; j<rqstdSeats_tmp[gender].length;j++){
		
    if(configRowOfSeat[prvsTmp[gender][i] ]  != configRowOfSeat[rqstdSeats_tmp[gender][j] ]  )continue;
		tmp=prvsTmp[gender][i]-rqstdSeats_tmp[gender][j];
		if (tmp <0 )tmp=-tmp;
		if (tmp==1){ sameRow=true; ; break};
		} // j
		if (sameRow)break; // no need to continue for thiis gender
		}  //i
		
		
		return sameRow;


}

   
//----------------------------------------------------------
function concatArrays(ar1,ar2){
var l1=ar1.length;
var jj;
for (jj=0;jj<ar2.length;jj++)ar1[l1+jj]=ar2[jj];
}



//----------------------------------------------------------	
	
 function convertListsToString(ar,idx){
  var rtrnVlu='';
	seperators=['$','+','-'];
	var i;
	var tmp;
	for (i=0;i<ar.length;i++){
	if (Array.isArray(ar[i])){ tmp=convertListsToString(ar[i],idx+1)}
	  else {	if (isNaN(ar[i]))  {tmp=ar[i] }
		    else tmp=ar[i].toString()};
	rtrnVlu=rtrnVlu+tmp+	seperators[idx];
	  }  //for		     
		rtrnVlu=rtrnVlu.substr(0,rtrnVlu.length-1);
		
		return 	rtrnVlu;
		}										 

////----------------------------------------------------------		 
	
	function evalRqstVSAssgnd(chag,row){
	// init default values of moedresults
	var i,j,k,m,amuda, gender,temp;
	var stsArry=[];
	var roww;
	var tempr=[];
	var counterNonAisles;
	var neededSeats,sum;
	var numOfRqstdCol = [[amudot.menRosh,amudot.womenRosh],[amudot.menKipur,amudot.womenKipur]];   //  [chag)[gender];
	
	var minNumberOfRows=[1,1];
	var numOfSeatsPerGenderPerChag=[[amudot.menRosh,amudot.menKipur],[amudot.womenRosh,amudot.womenKipur]];
	
	  	
/* 
calculate satisfaction 	for last year for a chag
     The process is: 
        First grade each assigned seat and then grade group of seats. 
      This is done separately for men and for women
	 
     Grade of a chair is 10 if in the requested list, priority 1; 9.5 if priority 2 and 9.1 if priority 3.
    Grade is 9 if not in the requested list but in the same row,  
        for men: 8 if adjacent row, 7 if 2 rows away, 6 for 3 rows away and  3 if in the same zone    and 0 if in a different ulam
For women: 8 if adjacent row, 6 if 2 rows away,   4 for 3 rows away and 3 if in the same zone and 0 if in a different ulam
     Grade of all chairs (men, women) is average of seats grade
	
   Factors for the group are calculated as follows:
Factor 1:
   If all the requested seats fit into n rows in the requested zone,    and they were allocated to m rows, then the factor is (n+2)/(m+2). 
The grade calculated in the previous step will be multiplied by this factor
Factor 2:
In the case where a few rows are designated as possible rows the program tries first to understand if each row is an alternative or all the designated seats are only one "alternative". To do that the program checks, in each row,  if enough seats are designated as "required" to accommodate the total request for the seats for this gender. If it is so,  the program decides that each row is an alternative.  In this case it checks if, excluding the "aisle" seats, there are still enough seats to accommodate the request. If there is at least one such row then it is understood that an aisle seat was not really a request.
In the case of "one alternative" the same decision process is applied for the "one alternative"
 
If aisle seat was requested  and there is no aisle seat allocated,   then the previous result will be multiplied by 0.85

All the above process is done per chag and the satisfaction grade per gender is the average of all the chagim (depending on  for how many chagim seats are requested)


All the above is repeated 3 times, first for priority1 seats, then priority 2 seats join the process and then priority 3.
Also here the program tries to determine if each priority is an alternative of the previous priority (if it is not adjacent to the previous seats) or an extension. If extension then priority 2 seats are added to the list of priority 1 seats and the process is repeated for the joined list, and if not the process is repeated for priority 2 only.
The same goes for priorities 2 and 3.


After repeating the process 3 times the program chooses the best result for each gender



 */ 
    individualGrades=[[],[]];  

   for (gender=0;gender<2;gender++){   // loop over gender
      assigns=assignedSeatslist[gender][chag];
			 if(dbgStsfction)console.log('gender='+gender+'   chag='+chag+'  assigns='+assigns);
			for (i=0;i<assigns.length;i++){
			   st=assigns[i];
				 
				 if (rqstdSeats[gender].indexOf(st) != -1) // assigned seat is in request list
				    { 
						   k= originalReqSeats.indexOf(st.toString());
			         k=Number(originalReqSeatPriority[k])-1;
				       kkkk=Math.pow( priorityFactorConst,k);
						   individualGrades[gender][i]=10*kkkk;
							 if(dbgStsfction)console.log('gggender='+gender+'  k='+k+'   kkkk='+kkkk+'  i='+i+'   individualGrades[gender][i]='+individualGrades[gender][i]);
							   continue;
						};
				temp=calcRowDistance(st,rqstdRows[gender][0]); 
				tempr=temp.split('$');
				
				k= originalReqSeats.indexOf(tempr[1]);
				if (k== -1){
				   kkkk=1;}   else{
			        k=Number(originalReqSeatPriority[k])-1;
				      kkkk=Math.pow( priorityFactorConst,k);
							}
				individualGrades[gender][i]=(10-tempr[0])*kkkk;
				if(dbgStsfction)console.log('gender='+gender+'  i='+i+'  k='+k+'   kkkk='+kkkk+'   individualGrades[gender][i]='+individualGrades[gender][i]);
			} // for i

	} // for gender			   		
			
			
			// calculate averages
			if ( ! individualGrades[0].length){menGrade=10} else menGrade=average(individualGrades[0]);
	   	if ( ! individualGrades[1].length){womenGrade=10} else womenGrade=average(individualGrades[1]);  
			
		
			//calculate and apply row factor --assgndRows[idx][chag]
		
				
				  for (gender=0;gender<2;gender++){   // loop over gender
					 amuda=numOfRqstdCol[chag][gender];						 
						neededSeats=Number(	requestedSeatsWorksheet[amuda+row].v)	;
						sum=0;
						eachRowIsAnAlternative=[true,true];
						for (k=0;k<rqstdRows[gender][1].length;k++)	{
						     sum=sum+	rqstdRows[gender][1][k];
								 if (sum <	neededSeats)minNumberOfRows[gender]++;
								 if (rqstdRows[gender][1][k] < neededSeats )eachRowIsAnAlternative[gender]=false;
								
								 }
						} // for gender		  
	    womenRowFactor=(minNumberOfRows[1]+2)/(assgndRows[1][chag].length+2); if (womenRowFactor> 1)womenRowFactor=1;
			menRowFactor=(minNumberOfRows[0]+2)/(assgndRows[0][chag].length+2);       if (menRowFactor > 1)menRowFactor=1;
			menGrade=menGrade*menRowFactor;
			womenGrade=womenGrade*womenRowFactor;
			
			
	 if(dbgStsfction)console.log('menRowFactor='+menRowFactor+' menGrade='+menGrade+'  womenRowFactor='+womenRowFactor+'   womenGrade='+womenGrade);
	 
	 var aisleFactor=[1,1];
	// calculate and apply aisle seats factor;
	    for (gender=0;gender<2;gender++){
				 if(dbgStsfction)console.log('gender='+gender+'  eachRowIsAnAlternative[gender]='+eachRowIsAnAlternative[gender]);

			 aisleReqstd=true;
			 ptr=numOfSeatsPerGenderPerChag[gender][chag]+row; 
			 rqstedSeatsPerChagPerGender=Number(requestedSeatsWorksheet[ptr].v);
			
			 if(eachRowIsAnAlternative[gender]){
			    
			    for (m=0 ;m<rqstdRows[gender][1].length; m++){   
					    counterNonAisles=rqstdRows[gender][1][m]-howManyAislesInRow(rqstdRows[gender][0][m],rqstdSeats[gender]);
							if (counterNonAisles  >= rqstedSeatsPerChagPerGender){  aisleReqstd=false; break;}
							 } // for m
					} // if 	eachRowIsAnAlternative	
				
				else {		
				 counterNonAisles=rqstdSeats[gender].length-howManyAislesInList(rqstdSeats[gender]);
				 if (rqstedSeatsPerChagPerGender <= counterNonAisles)aisleReqstd=false;
				 } // else
				 
				 aisleAssgnd=howManyAislesInList(assignedSeatslist[gender][chag]);
				 if(rqstdSeats[gender].length)  if (aisleReqstd && ( ! aisleAssgnd) )aisleFactor[gender]=0.85;
				  }
				 
				 
			menGrade=menGrade*aisleFactor[0];
			womenGrade=womenGrade*aisleFactor[1];

		 if(dbgStsfction)console.log('aisleFactor[0]='+aisleFactor[0]+' menGrade='+menGrade+'  aisleFactor[1]='+aisleFactor[1]+'   womenGrade='+womenGrade);

	//  
	
				 
			aChagRslts=[menGrade,womenGrade,combinedGrade]; 
						
			return;
}
//-----------------------------------------------------------------
function howManyAislesInRow(row, seatList){
  var i, roww,ptr,seat,tmp11;
	var list=[];
	var count=0;
	ptr=amudotOfConfig.open_badSeats	+  row.toString();
	
	//tmp11=  (shulConfigerationWS[ptr].v).toString();
	tmp=delLeadingBlnks(shulConfigerationWS[ptr].v); 
//	tmp=delLeadingBlnks(tmp11);
	if  ( !tmp ) return 0;
	list=tmp.split('+'); 
	for (i=0; i< seatList.length;i++){
		   seat=seatList[i];
	     roww=configRowOfSeat[Number(seat)];  
	     if (row == roww){
			   if (list.indexOf(seat.toString()) != -1)count++;
				 }
	} // for
	
		return count;
		
}		 
		

//-----------------------------------------------------------------

function howManyAislesInList(ar){
  var i,k;
	var seat;
	var tmp;
	var roww;
	var ptr;
	var list=[];
	var counter=0;
	
	 for (i=0; i<ar.length; i++){ 
	  seat=ar[i];  
		roww=configRowOfSeat[seat];
		ptr=amudotOfConfig.open_badSeats	+  roww.toString();
		tmp=delLeadingBlnks(shulConfigerationWS[ptr].v);
		if  ( !tmp ) continue;
		list=tmp.split('+'); 
		seat=seat.toString(); 
		for (k=0; k<list.length;k++)if (list[k]==seat) counter++; 
		}// for
		
		return counter;
	}

//-----------------------------------------------------------------

function average(ar){
 var sum;
 var i;
 sum=0;
 for (i=0;i<ar.length;i++)sum+=ar[i];
 return sum/ar.length;
 }
//-----------------------------------------------------------------


function calcRowDistance(seat,list){
  var i;
	var tmp;
	var row;
	var aNearRow;
	var dist;
	
	var ezor;
	var ulam;
	var ptr;
	var sameEzor=true;
  var sameUlam=true;
	
	ulam=UlamMartef[seat];
	ezor=ezorOfSeat[seat];
	dist=10000; // huge distance
	// find the nearest requested row
	row=configRowOfSeat[seat];
	foundInSameEzor=false;
	sameUlam=false;
	
	if ( ! list.length) return '0$1000';  // no request made; full satisfaction
	
	for (i=0;i <list.length;i++){
			
	  ptr=amudotOfConfig.fromSeat+(list[i]).toString();
    aSeatinThisRow=shulConfigerationWS[ptr].v;
							
	
	   tmp=row-list[i];
		 if (tmp <0)tmp=-tmp;
		 if (tmp < dist){ // found a near row candidate
		      if (ezorOfSeat[aSeatinThisRow] == ezor){
					   	dist=tmp;
							aNearRow=list[i];
							rslt=aSeatinThisRow.toString();
							foundInSameEzor=true;
							sameUlam=true;
							}  // if == ezor
			   if ( ! foundInSameEzor){ 
			       		if (! sameUlam)	sameUlam=(UlamMartef[aSeatinThisRow] == ulam);
								}
								
					}  // if tmp< dist
			}  // for i						
							   
					 
		if ( ! sameUlam ) { return '10$1000';};   // deduct 10 from max grade  (from 10) ; row that does not exist
		
		if ( ! foundInSameEzor ) { return '7$1000';};   // deduct 7 from max grade  (from 10)  ; row that does not exist
		
		// in the same ezor
		
			   ptr1=amudotOfConfig.reltvRowQual+aNearRow.toString();
				 ptr2=amudotOfConfig.reltvRowQual+row.toString();
         delta=shulConfigerationWS[ptr1].v-shulConfigerationWS[ptr2].v;

				if (badSeats.indexOf(seat) != -1) delta++;  // 
			    if (delta<0)delta=dist;
					if (dist<delta)dist=delta;
					if (dist> 6)	dist=6;
					if( dist == 0)dist=1;  // same line
					rslt=dist.toString()+'$'+rslt;
					
					return rslt;
						
}

/*----------------------------------------------------------
function moveArrays(target,source){ return; 
  var i;
	for (i=0; i<source.length;i++){
	  if (Array.isArray(source[i])){moveArrays(target[i],source[i]) }
		    else  {  if (source[i]== source[i]+0){target[i]=source[i]+0;} else target[i] =source[i]+'';  //add nothing to value
								
						 }  // else
	}  //for
	return
	}			
*/

//----------------------------------------------------------
app.get('/testStsfction', function(req, res) {
	res.header("Access-Control-Allow-Origin", "*");
	 res.setHeader('Content-Type', 'text/html');
	
	 
	 inputString=decodeURI(req.originalUrl).split('?')[1];
	 inputParams=inputString.split('$'); 
	 chagdbg=Number(inputParams[1]);
	 rowdbg=Number(inputParams[2]);
	 if (inputParams[0] == mngmntPASSW){
	    dbgStsfction=true;
			initFromFiles('2016');
	    analyseRqstVSAssgnd(rowdbg);
      dbgStsfction=false;
			}
res.send('---' );

   })
//----------------------------------------------------------



app.get('/updateAssignedSeats', function(req, res) {
	res.header("Access-Control-Allow-Origin", "*");
	 res.setHeader('Content-Type', 'text/html');
	
	 
	 inputString=decodeURI(req.originalUrl).split('?')[1];
	 inputParams=inputString.split('$'); 
	 if (inputParams[0] != mngmntPASSW){res.send('999' ); return};
	 
	 initFromFiles('');
	   moed=inputParams[1]; 
		 moed_integer=[' ','rosh','kipur'].indexOf(moed);
	   nameToUpdate=inputParams[2];  
	   strSeatsToUpdate=inputParams[3];  
	 
	   rowNum= knownName(nameToUpdate)[0];   
		 if(rowNum == -1 ){res.send('---' ); return};
		 row=rowNum.toString();
		 
		 errInAsgnmnt=false;
		 if( (moed=='rosh') || (moed=='all') )errInAsgnmnt =checkRequestedAssgnmntForDoubles(strSeatsToUpdate, 1,rowNum);  // function returns true if a problem IS found
		  if( (moed=='kipur') || (moed=='all') )errInAsgnmnt =errInAsgnmnt || checkRequestedAssgnmntForDoubles(strSeatsToUpdate, 2,rowNum);  // function returns true if a problem IS found
			
			if (errInAsgnmnt){res.send('???'); return}; //found doubles
			
		 if( (moed=='rosh') || (moed=='all') ){
		 assgndBegunRosh=true;
		 ptr=amudot.assignedSeatsRosh+row; 
		 oldAssignedSeatsRosh=requestedSeatsWorksheet[ptr].v;
		 requestedSeatsWorksheet[ptr].v=strSeatsToUpdate; 
		  
		 closeSeats(1,row);
		 markedSeatsLeft('rosh',row,delLeadingBlnks(oldAssignedSeatsRosh)); 
			 };
			 
		 if( (moed=='kipur') || (moed=='all') ){ 
		 assgndBegunKipur=true;  
		 ptr=amudot.assignedSeatsKipur+row;
		 oldAssignedSeatsKipur=requestedSeatsWorksheet[ptr].v;
		 requestedSeatsWorksheet[ptr].v=strSeatsToUpdate;   
		 closeSeats(2,row);
		 markedSeatsLeft('kipur',row,delLeadingBlnks(oldAssignedSeatsKipur));
		    }
		 CountAssignedPerMoed_PerUlam();
			 xlsx.writeFile(workbook, XLSXfilename);
			 
			 if (
			     (requestedSeatsWorksheet[amudot.numberOfAssignedSeatsRoshMen+row].v == Number(requestedSeatsWorksheet[amudot.menRosh+row].v)   )
					 && (requestedSeatsWorksheet[amudot.numberOfAssignedSeatsRoshWomen+row].v == Number(requestedSeatsWorksheet[amudot.womenRosh+row].v ) )
					 && (requestedSeatsWorksheet[amudot.numberOfAssignedSeatsKipurMen+row].v == Number(requestedSeatsWorksheet[amudot.menKipur+row].v))
					 && (requestedSeatsWorksheet[amudot.numberOfAssignedSeatsKipurWomen+row].v == Number(requestedSeatsWorksheet[amudot.womenKipur+row].v ) )
					 ){
					   
					   msg=calculate_crnt_assnmnt_stsfctn(row);
						 res.send('***%'+msg)
						 }
				else {
				      requestedSeatsWorksheet[amudot.stsfctnInFlrThisYrMen+row].v=' ';
	            requestedSeatsWorksheet[amudot.stsfctnInFlrThisYrWmn+row].v=' ';
	            requestedSeatsWorksheet[amudot.ThisYRSSeat+row].v=' ';
				      res.send('+++' );
			      };
			
			} )
	 
	 
//----------------------------------------------------------      
app.get('/reCalcStsfction', function(req, res) {
	res.header("Access-Control-Allow-Origin", "*");
	 res.setHeader('Content-Type', 'text/html');
	
	 var str;
	 inputString=decodeURI(req.originalUrl).split('?')[1];
	 inputParams=inputString.split('$'); 
	 if (inputParams[0] != debugPASSW){res.send('999' ); return};
	 
	 initFromFiles('');
  row=inputParams[1];
	 str=calculate_crnt_assnmnt_stsfctn(row);
	 res.send('+++'+str);
	 
	 
	 	} )
//----------------------------------------------------------
							
/*

 amudotForStsfctn=[	amudot.stsfctnInFlrThisYrWmn,amudot.stsfctnInFlrThisYrMen,amudot.ThisYRSSeat,
							          amudot.stsfctnInFlrLastYrWmn,amudot.stsfctnInFlrLastYrMen,amudot.lstYrSeat,
							          amudot.stsfctnInFlr2YRSAgoYrWmn,amudot.stsfctnInFlr2YRSAgoYrMen,amudot.TwoYRSAgoSeat,
												amudot.stsfctnInFlr3YRSAgoYrWmn,amudot.stsfctnInFlr3YRSAgoYrMen,amudot.ThreeYRSAgoSeat];	
*/												
												
function calculate_crnt_assnmnt_stsfctn(row){
var tmp, str,i,colmn;
  
	
	stsfctionColction=[];
	analyseRqstVSAssgnd(row);
		
		
	tmp=stsfctionColction[0].split('$');
	 // console.log('tmp='+tmp);
  requestedSeatsWorksheet[amudot.stsfctnInFlrThisYrMen+row].v=tmp[1]+'*';;
	requestedSeatsWorksheet[amudot.stsfctnInFlrThisYrWmn+row].v=tmp[2]+'*';
	requestedSeatsWorksheet[amudot.ThisYRSSeat+row].v=tmp[3]+'*';
	
		
		xlsx.writeFile(workbook, XLSXfilename);		
	  str='';
		for (i=0;i<amudotForStsfctn.length;i++){
		    colmn=amudotForStsfctn[i];
				str=str+requestedSeatsWorksheet[colmn+row].v+'$';
				}
				str=str+requestedSeatsWorksheet[amudot.memberShipStatus+row].v;
			//	console.log('calculate_crnt_assnmnt_stsfctn     str='+str);
				return str;
}		
	
//----------------------------------------------------------
 function updateRowForNewSelection(roww){
 
  requestedSeatsWorksheet[amudot.notAssignedMarkedSeatsKipur+roww].v=requestedSeatsWorksheet[amudot.markedSeats+roww].v;
 requestedSeatsWorksheet[amudot.notAssignedMarkedSeatsRosh+roww].v=requestedSeatsWorksheet[amudot.markedSeats+roww].v;
 //requestedSeatsWorksheet[amudot.assignedSeatsKipur+roww].v='';
 //requestedSeatsWorksheet[amudot.assignedSeatsRosh+roww].v='';
 //requestedSeatsWorksheet[amudot.numberOfAssignedSeatsRoshMen+roww].v='0';
// requestedSeatsWorksheet[amudot.numberOfAssignedSeatsRoshWomen+roww].v='0';
 //requestedSeatsWorksheet[amudot.numberOfAssignedSeatsKipurMen+roww].v='0';
// requestedSeatsWorksheet[amudot.numberOfAssignedSeatsKipurWomen+roww].v='0';

if(assgndBegunRosh )init_notAssigenedMarked('rosh');
if(assgndBegunKipur )init_notAssigenedMarked('kipur');

 wmn=0;
 men=0;
 sts=delLeadingBlnks(requestedSeatsWorksheet[amudot.markedSeats+roww].v);
 
 if (sts) {
   stsArry=sts.split('+');
   for(iill=0;iill<stsArry.length;iill++){
     st=Number(stsArry[iill].split('_')[0]);
		 if( isWoman[st] ){ wmn++; }  else men++;
		} 
	}	
		wmn=wmn.toString();
		men=men.toString(); 
 requestedSeatsWorksheet[amudot.numberMarkedMen+roww].v=men;
 requestedSeatsWorksheet[amudot.numberMarkedWomen+roww].v=wmn;
 requestedSeatsWorksheet[amudot.NumberOfNotAssignedMarkedSeatsMen+roww].v=men;
 requestedSeatsWorksheet[amudot.NumberOfNotAssignedMarkedSeatsWomen+roww].v=wmn;


}
//----------------------------------------------------------
function init_notAssigenedMarked(moed){
 assgndBegunRosh=false;
 assgndBegunKipur=false;
 
 if(moed=='rosh'){
     assgnCol=amudot.assignedSeatsRosh;
		 stillRequestedForMoedCol=amudot.notAssignedMarkedSeatsRosh;
	
		 markedLeftCol=amudot.notAssignedMarkedSeatsRosh;
		 namesAssignedIdx=0;
		          }
		else {
		  assgnCol=amudot.assignedSeatsKipur;
			stillRequestedForMoedCol=amudot.notAssignedMarkedSeatsKipur;
			markedLeftCol=amudot.notAssignedMarkedSeatsKipur;
			namesAssignedIdx=1;
			  }

 for (membr1=firstSeatRow; membr1<lastSeatRow+1; membr1++){
      requestedSeatsWorksheet[stillRequestedForMoedCol+membr1.toString()].v=requestedSeatsWorksheet[amudot.markedSeats+membr1.toString()].v;
			if(delLeadingBlnks(requestedSeatsWorksheet[amudot.assignedSeatsRosh +row].v) )assgndBegunRosh=true;
			if(delLeadingBlnks(requestedSeatsWorksheet[amudot.assignedSeatsKipur +row].v) )assgndBegunKipur=true;

			};
				
	for (rownum=firstSeatRow; rownum<lastSeatRow+1; rownum++){
	 row=rownum.toString();			
				
	 if( ! delLeadingBlnks(requestedSeatsWorksheet[assgnCol +row].v))continue;
	 newlyAssignedArray=(requestedSeatsWorksheet[assgnCol +row].v).split('+');
	          
					 for (membr1=firstSeatRow; membr1<lastSeatRow+1; membr1++){
								   row1=membr1.toString();  
									
											LeftMarkedSeatsArray=(requestedSeatsWorksheet[stillRequestedForMoedCol +row1].v).split('+');
											
                         for(ii=0; ii<newlyAssignedArray.length; ii++){
												      for (ij=0; ij < LeftMarkedSeatsArray.length; ij++)
															          if(LeftMarkedSeatsArray[ij].split('_')[0] == newlyAssignedArray[ii]){
																				
																				                      LeftMarkedSeatsArray.splice(ij,1);
																													//		break;   
																															}  // for ij
														            															
												      		     }  // for ii
                    	requestedSeatsWorksheet[stillRequestedForMoedCol +row1].v=LeftMarkedSeatsArray.join('+');															

													}  // for membr1	
										} // for rownum					
						}

//----------------------------------------------------------

function markedSeatsLeft(moed,row,oldAssignment){
var nm;
var tmp;
 if(moed=='rosh'){
     assgnCol=amudot.assignedSeatsRosh;
		 stillRequestedForMoedCol=amudot.notAssignedMarkedSeatsRosh;
		 markedLeftCol=amudot.notAssignedMarkedSeatsRosh;
		 namesAssignedIdx=0;
		          }
		else {
		  assgnCol=amudot.assignedSeatsKipur;
			stillRequestedForMoedCol=amudot.notAssignedMarkedSeatsKipur;
			markedLeftCol=amudot.notAssignedMarkedSeatsKipur;
			namesAssignedIdx=1;
			  }


	 if(requestedSeatsWorksheet[assgnCol +row].v){newlyAssignedArray=(requestedSeatsWorksheet[assgnCol +row].v).split('+') }
	           else newlyAssignedArray=[];
								 for (membr1=firstSeatRow; membr1<lastSeatRow+1; membr1++){
								      row1=membr1.toString();
											LeftMarkedSeatsArray=(requestedSeatsWorksheet[stillRequestedForMoedCol +row1].v).split('+');
                         for(ii=0; ii<newlyAssignedArray.length; ii++){
												      for (ij=0; ij < LeftMarkedSeatsArray.length; ij++)
															          if(LeftMarkedSeatsArray[ij].split('_')[0] == newlyAssignedArray[ii]){
																				                      LeftMarkedSeatsArray.splice(ij,1);
																													//		break;   
																															}  // for ij
														            	requestedSeatsWorksheet[stillRequestedForMoedCol +row1].v=LeftMarkedSeatsArray.join('+');															
												      		     }  // for ii
															}  // for membr1	
															
														
	// update namesForSeat with new assigns
		for (ii=0; ii<		newlyAssignedArray.length;ii++){
		   			aSeat=Number(newlyAssignedArray[ii].split('_')[0]);
						namesForSeatParts=namesForSeat[aSeat].split('/');
            namesAssigned=namesForSeatParts[0].split('$');
						nm=delLeadingBlnks( requestedSeatsWorksheet[amudot.name +row].v);
            tmp=nm.substr(nm.length-1);
	          if(tmp=='*')nm=nm.substr(0,nm.length-1);
			      namesAssigned[namesAssignedIdx]=nm;
			      namesForSeatParts[0]=namesAssigned.join('$');
	          namesForSeat[aSeat]=namesForSeatParts.join('/'); 
						}								
															
  // create listof un-assigned seats
	 if ( ! oldAssignment) return; // no old assignment
	 
	
	  oldAssignedArray=	oldAssignment.split('+');
		unAssigned_array=[];
		for (ii=0;  ii	<		oldAssignedArray.length;	ii++)
		   if(newlyAssignedArray.indexOf(oldAssignedArray[ii]) == -1)unAssigned_array.push(oldAssignedArray[ii]); // this seat is not assigned anymore to this member
		
		for (ii=0; ii <unAssigned_array.length; ii++){
		   aSeat=Number(unAssigned_array[ii].split('_')[0]); 
			 
			 // delete un assigned from namesForSeat
			 namesForSeatParts=namesForSeat[aSeat].split('/');
       namesAssigned=namesForSeatParts[0].split('$');
			 namesAssigned[namesAssignedIdx]='';
			 namesForSeatParts[0]=namesAssigned.join('$');
	     namesForSeat[aSeat]=namesForSeatParts.join('/');
			 
			 aSeatSTR=aSeat.toString();
			 requestorsList= (namesForSeat[aSeat].split('/')[1]).split('$');  
			 for (ij=0; ij<requestorsList.length; ij++){
			   member=  requestorsList[ij]; 
			   memberRow=knownName(member)[0];
				 if( memberRow == -1) continue;
				 row1=memberRow.toString();
				 marked=delLeadingBlnks(requestedSeatsWorksheet[amudot.markedSeats +row1].v);
				  if ( ! marked) continue; // no old request
				 markedList=marked.split('+');
				 wasInMarkedList=false;
				 for (ik=0; ik<markedList.length; ik++)
				                           if(markedList[ik].split('_')[0] == aSeatSTR){ wasInMarkedList=true; break;};
				if ( !	wasInMarkedList) continue;	
				markedLeftSTR=delLeadingBlnks(requestedSeatsWorksheet[markedLeftCol +row1].v);
				markedLeftList=	markedLeftSTR.split('+');
				notFound=true;
				for(ik=0; ik<markedLeftList.length;ik++)if (markedLeftList[ik].split('_')[0]==aSeatSTR)notFound=false;
				if(notFound){
				      markedLeftList.push	(aSeatSTR);
							markedLeftList=markedLeftList.sort(sortOrder);
							requestedSeatsWorksheet[markedLeftCol +row1].v=markedLeftList.join('+');							 
			  } 
		}// for ij
		} //for ii	 																						
}		
	
//---------------------------------------------------------- 
app.get('/setDebugOn', function(req, res) {
	 res.header("Access-Control-Allow-Origin", "*");
	 res.setHeader('Content-Type', 'text/html');
	 
	 
	 var i;
	 
	 inputString=decodeURI(req.originalUrl).split('?')[1];   
	 debugparam=inputString.split('$'); 
	 if  ( debugparam[0] != debugPASSW){ res.send('wrong password'); return};
	 
	 if(debugparam[1]=='query'){
	      res.send('ok %'+debugRequestsAll.join('*'));
				return};
	  console.log(' debugparam='+debugparam);  
		 
	 if(debugparam[1]=='reset'){
	   
	          debugRequestsKeys=[];
						debugRequestsAll=[];
				}
					
	   
		    
	
	
	 // if request is already there - remove it.(not to have the same request twice and to make sure that current params are kept
	 i=searchDebugParam(debugparam[2]);
		 if (i != -1){
		       debugRequestsKeys.splice(i,1);
	         debugRequestsAll.splice(i,1);
					 }
	     if(debugparam[1]=='on'){  // if a request is on put it in
	     // console.log('setdebug length='+debugRequestsKeys.length);
				if (debugRequestsKeys.length > 20 ){ res.send('too  many debug commands'); return};
	      debugRequestsKeys[debugRequestsKeys.length]=debugparam[2];
			  debugRequestsAll[debugRequestsAll.length]=inputString;
		}
	 
	  if(debugparam[1]=='off'){  // already removed
		// do nothing
		 };
		
	for (i=0;i<debugRequestsKeys.length;i++){	
		 ptrA=amudot.debugRequestsKeys+(i+3).toString(); // first row is 3
	   requestedSeatsWorksheet[ptrA].v=debugRequestsKeys[i];
		 ptrB=amudot.debugRequestsAll+(i+3).toString(); 
	   requestedSeatsWorksheet[ptrB].v=debugRequestsAll[i];
		
		 };
		 
		 
		 for (j=i;j<20;j++){
		   ptrA=amudot.debugRequestsKeys+(j+3).toString(); // first row is 3
	     requestedSeatsWorksheet[ptrA].v='$$$';  // $$$ == empty slot
		   ptrB=amudot.debugRequestsAll+(j+3).toString(); 
	     requestedSeatsWorksheet[ptrB].v='$$$';
				}
			xlsx.writeFile(workbook, XLSXfilename);	
	res.send('ok %'+debugRequestsAll.join('*'));
  
        })	
				
//---------------------------------------------------------------------------------	
function searchDebugParam(param){
   var i;  
	 for (i=0; i<	debugRequestsKeys.length;i++)if (		debugRequestsKeys[i] == param )return i;
	 return -1;
	 }
//---------------------------------------------------------------------------------	
app.get('/getRowValues', function(req, res) {
	res.header("Access-Control-Allow-Origin", "*");
	 res.setHeader('Content-Type', 'text/html');
	 inputString=decodeURI(req.originalUrl).split('?')[1];
	
	 inputPrms=inputString.split('$');  //  name $   passw   $   req_year  $
	 nameToDebug=[];
	 if ( inputPrms[1]==debugPASSW){
	     req_year=inputPrms[2];
			 initFromFiles(req_year);
	     nameToDebug=knownName(inputPrms[0]);  
			 numberOfFirstnames=nameToDebug[1].split('$').length;
			 if (numberOfFirstnames > 1) {
			     res.send('999 name not well defined') }
					else {
					   if (nameToDebug[0] == -1) {
						     res.send('888 name does not exist') }
								   else {  // name exists and isn unique
									    listToSend='';
											roww=nameToDebug[0].toString();
											 Object.keys(amudot).forEach(function(key)  {   // copy all hdrs and values for row
											    colmn=amudot[key];
													if (   requestedSeatsWorksheet[ colmn+roww] )
												      listToSend=listToSend+colmn+'&'   //key
													      + requestedSeatsWorksheet[ colmn+'1'].v+'&'  //hdr
													      + requestedSeatsWorksheet[ colmn+roww].v     //value 
												       	+'$';                                                  // delimiter
													 }) // for each 
											listToSend=listToSend.substr(0,listToSend.length-1);		 
											
											res.send('000$'+listToSend);
			              }  // else  name exists
         } // else
				} else  res.send('777 wrong debug password')
			})								
											
											
//---------------------------------------------------------------------------------	
app.get('/manualUpdateValues', function(req, res) {
	res.header("Access-Control-Allow-Origin", "*");
	 res.setHeader('Content-Type', 'text/html');
	 inputString=decodeURI(req.originalUrl).split('?')[1];
	
	 inputPrms=inputString.split('^');
   if (inputPrms[0] != debugPASSW){
	   res.send('777 wrong debug password');
		 return;
		 }
		
		inputPrms=inputPrms[1].split('$');  
		temp1=inputPrms[0].split('&'); 
		nam=temp1[1];
		if (nam.substr(nam.length-1,1) =='*')nam=nam.substr(0,nam.length-1);
		    		
	  rowToDebug=knownName(nam)[0]; 
		if (rowToDebug == -1){
		  res.send('888 name does not exist');
			return; }
	
		rowToDebug=rowToDebug.toString();	
		for(i=0;i<inputPrms.length;i++){
		  temp1=inputPrms[i].split('&');
			ptr=temp1[0]+rowToDebug;
			vlu=temp1[2];
		//console.log('i='+i+' ptr='+ptr+'   vlu='+vlu);
			
			requestedSeatsWorksheet[ptr].v=vlu;
			}  //for
			
			xlsx.writeFile(workbook, XLSXfilename);
			res.send('000 updated');

})

//--------------------------------------------------------------------------------
		
app.get('/getMembersInfo', function(req, res) {
	res.header("Access-Control-Allow-Origin", "*");
	inputString=decodeURI(req.originalUrl).split('?')[1]; 
	if (inputString != mngmntPASSW){
	   res.send('--- wrongg  password');
	 }
		 initFromFiles('');  // last year info  
	listToSend='+++$';
	 Object.keys(amudotForMemberInfo).forEach(function(key)  {   // copy all hdrs 
											    colmn=amudot[key];
													    listToSend=listToSend+delLeadingBlnks(requestedSeatsWorksheet[ colmn+'1'].v)+'&';  //hdrs
													      
													 }) // for each
													  listToSend=listToSend.substr(0,listToSend.length-1); //remove last & 
											  
	  for (member=firstSeatRow; member<lastSeatRow+1; member++){ 
		temp1=listToSend.length;   
		    listToSend=listToSend+'$'		
		    sMember=member.toString(); 
				 Object.keys(amudotForMemberInfo).forEach(function(key)  {   // copy all hdrs 
											    colmn=amudot[key];
													ptr=colmn+sMember;
													vlu='';   if ( requestedSeatsWorksheet[ ptr] ) vlu=requestedSeatsWorksheet[ ptr].v;
													if ( ! isNaN(vlu)  )vlu=vlu.toString();  vlu=delLeadingBlnks(vlu);
													    listToSend=listToSend+vlu+'&';  
													      
													 }) // for each 
					listToSend=listToSend.substr(0,listToSend.length-1); //remove last & 
					temp2=listToSend.length;   
					toPrint=listToSend.substr(temp1,temp2-temp1);
				
					} // for member								 
	   res.send(listToSend);
 });	


		
//---------------------------------------------------------------------------------	  
app.get('/getListOfRegistered', function(req, res) {
	res.header("Access-Control-Allow-Origin", "*");
   inputString=decodeURI(req.originalUrl); 
  info=inputString.split('-')[1];
	info=info.split('$');
	passW=info[0];
	DST_inIsrael=Number(info[1]);
	 if (passW == mngmntPASSW){  fileToSendName= 'registerdList.xlsx';  fileToRead=registeredListFilename; generate_registeredList_XLS(DST_inIsrael);}
				else {fileToSendName='empty.xlsx'; fileToRead='empty.xlsx'};
	 
     res.setHeader('Content-type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
        res.setHeader('Content-Disposition', 'attachment; filename=' + fileToSendName);
	      var fileR = fs.readFileSync(fileToRead, 'binary');
        res.setHeader('Content-Length', fileR.length);
        res.write(fileR, 'binary');
        res.end();
      
 
});
//---------------------------------------------------------------------------------	 
 

// initialize membersRequests.xlsx file

app.get('/s276662', function(req, res) {
	res.header("Access-Control-Allow-Origin", "*");
	tmpfile=fs.readFileSync('membersRequests.xlsx');
	fs.writeFileSync(XLSXfilename, tmpfile);
	workbook = xlsx.readFile(XLSXfilename);
	requestedSeatsWorksheet = workbook.Sheets['HTMLRequests'];  
	
	tmpfile=fs.readFileSync('supportTables.xlsx');
	fs.writeFileSync(supportTblsFilename, tmpfile);
	supportWB=xlsx.readFile(supportTblsFilename); 
	
	initFromFiles('');
	
	
	 res.setHeader('Content-Type', 'text/html'); 
	res.send(cache_get('initialize') );
	
	 
	 });
//----------------------------------------------------------
app.get('/getYearList', function(req, res) {
	res.header("Access-Control-Allow-Origin", "*");
	 res.setHeader('Content-Type', 'text/html');
	
	 years=[];   
	 shNames=workbook.SheetNames;  
	 for (i=0; i<shNames.length;i++)  
	       if(shNames[i].substr(0,12)  == 'HTMLRequests')
				     if( shNames[i].substr(12) ) years.push(shNames[i].substr(12)); // not null
				 // for	
				
			  
	 res.send(years.join('$'));
})




//---------------------------------------------------------- 

app.get('/setreadxls', function(req, res) {
	res.header("Access-Control-Allow-Origin", "*");
	 res.setHeader('Content-Type', 'text/html');
	 initFromFiles('');
	 inputString=decodeURI(req.originalUrl).split('?')[1];
	 inpArray=inputString.split('$');
	 if (inpArray[0] == '987882'){
	      if ( inpArray[2]  ) { // set value
				   requestedSeatsWorksheet[inpArray[1]].v=inpArray[2];
					 xlsx.writeFile(workbook, XLSXfilename);
					 }
		    vlu=	requestedSeatsWorksheet[inpArray[1]].v	;	
				res.send('value in cell '+ inpArray[1]+ ' is '+vlu);
			} 
		else		res.send('err in passw');
})
//------------------------------------------------------------------------------	
  app.get('/roshToKipur', function(req, res) {
	res.header("Access-Control-Allow-Origin", "*");
	 res.setHeader('Content-Type', 'text/html');
	 initFromFiles('');
   var tmp,i,row;
	 var notCopied=[];
	 
	 for (i=firstSeatRow;i<lastSeatRow+1;i++){  
	       tmp=delLeadingBlnks(requestedSeatsWorksheet[amudot.assignedSeatsRosh +row].v);
	       row=i.toString();
			   if(	 checkRequestedAssgnmntForDoubles(tmp,2,i)  ){          //(reqstedAssgmnt, moed,checkedRow);; value true if problem
				   notCopied.push(row);} else  
					       requestedSeatsWorksheet[amudot.assignedSeatsKipur +row].v=requestedSeatsWorksheet[amudot.assignedSeatsRosh +row].v
				
				 }
				 
     xlsx.writeFile(workbook, XLSXfilename);

    initValuesOutOfHtmlRequestsXLSX_file();   //init values

    init_notAssigenedMarked('rosh');
	  init_notAssigenedMarked('kipur');

    CountAssignedPerMoed_PerUlam();
		
		checkDoubeeAssignments();
      tmp=notCopied.join('+');
			res.send('copied$'+tmp);	 
	}) 
//------------------------------------------------------------------------------						
	
	app.get('/initNextyearFiles', function(req, res) {
	res.header("Access-Control-Allow-Origin", "*");
	 res.setHeader('Content-Type', 'text/html');
	
	 inputString=decodeURI(req.originalUrl).split('?')[1];  
	 msgParts=inputString.split('$');
	 if (msgParts[0] != mngmntPASSW){console.log('wrong password'); res.send('999' )}
	 else {
	 
	    newYear=msgParts[1];
  		
			resltSupportTbl=genNewYearSpportTblSheet(newYear);		
       resltREQ=genNewYearRequestSheet(newYear); 
        res.send(resltREQ+'$'+resltSupportTbl);
      			
							
							
							
									
			}  // else						
								 
 
	 })
	 
//---------------------------------------------------------- 
function genNewYearRequestSheet(yearToCreate){

		 YearToSaveSheetName=	'HTMLRequests'	+ 	(yearToCreate-1).toString();							
	       nm=workbook.SheetNames.length; 
         sheetNum=workbook.SheetNames.indexOf(YearToSaveSheetName); 
			
         if(sheetNum != -1) return '900'; // requested saving already done
				
				
				workbook.SheetNames[nm]=YearToSaveSheetName; 
								
								newWs = creatSheet(numOfColsInNewSheet ,  numOfRowsInNewSheet);  
								workbook.Sheets[YearToSaveSheetName]=newWs;
						
								YrSheet=workbook.Sheets[YearToSaveSheetName]; 
								
								for (rw=1; rw<numOfRowsInNewSheet;rw++) {      // copy all rows to file to Save a 
								       rww=rw.toString();
								       Object.keys(amudot).forEach(function(key)  {   // copy all colomns

                            ptr1= amudot[key]+rww;
														vlu=' ';
														if( requestedSeatsWorksheet[ptr1]) vlu=requestedSeatsWorksheet[ptr1].v;
									        //console.log('ptr1='+ptr1+'   vlu='+vlu);
														YrSheet[ptr1]={t:"s",v:vlu};
														YrSheet[ptr1].v=(YrSheet[ptr1].v).toString();  //make sure it is of str attribute
												
                         });	
				if (rw <  3 )continue; // header lines
				
				if ( ! requestedSeatsWorksheet[amudot.name+rww]) continue;
				
				tmp=delLeadingBlnks(requestedSeatsWorksheet[amudot.name+rww].v); 
				if ( ( ! tmp) || (tmp=='$$$')  )continue;  //empty row
												 
				         // move stsfction one year back
								 
	 requestedSeatsWorksheet[amudot.stsfctnInFlr3YRSAgoYrWmn+rww]=requestedSeatsWorksheet[amudot.stsfctnInFlr2YRSAgoYrWmn+rww];
	 requestedSeatsWorksheet[amudot.stsfctnInFlr3YRSAgoYrMen+rww]=requestedSeatsWorksheet[amudot.stsfctnInFlr2YRSAgoYrMen+rww];
	 requestedSeatsWorksheet[amudot.ThreeYRSAgoSeat+rww]=requestedSeatsWorksheet[amudot.TwoYRSAgoSeat+rww];
	 
	 requestedSeatsWorksheet[amudot.stsfctnInFlr2YRSAgoYrWmn+rww]=requestedSeatsWorksheet[amudot.stsfctnInFlrLastYrWmn+rww];
	 requestedSeatsWorksheet[amudot.stsfctnInFlr2YRSAgoYrMen+rww]=requestedSeatsWorksheet[amudot.stsfctnInFlrLastYrMen+rww];
	 requestedSeatsWorksheet[amudot.TwoYRSAgoSeat+rww]=requestedSeatsWorksheet[amudot.lstYrSeat+rww];
	 
	 requestedSeatsWorksheet[amudot.stsfctnInFlrLastYrWmn+rww]=requestedSeatsWorksheet[amudot.stsfctnInFlrThisYrWmn+rww];
	 requestedSeatsWorksheet[amudot.stsfctnInFlrLastYrMen+rww]=requestedSeatsWorksheet[amudot.stsfctnInFlrThisYrMen+rww];
	 requestedSeatsWorksheet[amudot.lstYrSeat+rww]=requestedSeatsWorksheet[amudot.ThisYRSSeat+rww];
	 
	 requestedSeatsWorksheet[amudot.stsfctnInFlrThisYrWmn+rww].v=' ';
	 requestedSeatsWorksheet[amudot.stsfctnInFlrThisYrMen+rww].v=' ';
	 requestedSeatsWorksheet[amudot.ThisYRSSeat+rww].v=' ';
	 
		
			
								  	// update membership status   
									ptr1=amudot.memberShipStatus+rww;
									vlu=Number(requestedSeatsWorksheet[ptr1].v);
									if(vlu<2)requestedSeatsWorksheet[ptr1].v=(vlu+1).toString();				 
								 
												 
									// clear last  year ibfo from requested seats worksheet for the new yr new info
                  Object.keys(amudotToClrInReqstdSeatsWhnGenNewYr).forEach(function(key)  {

                            ptr1= amudotToClrInReqstdSeatsWhnGenNewYr[key]+rww;
									        
													requestedSeatsWorksheet[ptr1]= {t:"s",v:' '};
													requestedSeatsWorksheet[ptr1].v=(requestedSeatsWorksheet[ptr1].v).toString();  //make sure it is of str attribute
													
														 
													 });				 
									}	// for
								requestedSeatsWorksheet['C2']= {t:"s",v:' '};	// clear closing date and time
								
								for (i=0;i<300;i++){   // init transaction history
								  ptr1='A'+(transactionHistoryFirstRow+i).toString();
									requestedSeatsWorksheet[ptr1].v='$$$';
							}		
							
							xlsx.writeFile(workbook, XLSXfilename);	
							workbook = xlsx.readFile(XLSXfilename);  

	 	
							
							initValuesOutOfHtmlRequestsXLSX_file('');
				
		return '+++';
									
	}
	

//---------------------------------------------------------- 
function genNewYearSpportTblSheet(yearToCreate){

// save isWoman sheet 

   rsltWmn='+++';
   shulConfigerationNumOfCols=9;        
	 shulConfigerationNumOfRows=800;
	 shulConfigerationAmudot=['A','B','C','D','E','F','G','H','I'];
		 YearToSaveShulConfigerationSheetName=	'shulConfigeration'	+ 	(yearToCreate-1).toString();							
	       nm=supportWB.SheetNames.length; 
         sheetNum=supportWB.SheetNames.indexOf(YearToSaveShulConfigerationSheetName); 
			   lastYrshulConfigerationSheet=supportWB.Sheets['shulConfigeration']; 
         if(sheetNum != -1){ rsltWmn= '900'; }     // requested saving already done
				 else {
				
				supportWB.SheetNames[nm]=YearToSaveShulConfigerationSheetName; 
								
								newWs = creatSheet(shulConfigerationNumOfCols ,  shulConfigerationNumOfRows);  
								supportWB.Sheets[YearToSaveShulConfigerationSheetName]=newWs;
						
								YrSheet=supportWB.Sheets[YearToSaveShulConfigerationSheetName]; 
								
								for (rw=1; rw<shulConfigerationNumOfRows;rw++) {      // copy all rows to file to Save a 
								       rww=rw.toString();
								       Object.keys(shulConfigerationAmudot).forEach(function(key)  {   // copy all colomns

                            ptr1= shulConfigerationAmudot[key]+rww;
														vlu=' ';
														if( lastYrshulConfigerationSheet[ptr1]) vlu=lastYrshulConfigerationSheet[ptr1].v;
									        
													 
														YrSheet[ptr1]={t:"s",v:vlu};
														YrSheet[ptr1].v=(YrSheet[ptr1].v).toString();  //make sure it is of str attribute
												
                         });	
								 
									}	// for
					} //else				
								
	// save seat to row sheet	
		rsltToRow='+++';
		 seatToRowNumOfCols=4;
	 seatToRowNumOfRows=1400;
	 seatToRowAmudot=['A','B','C','D'];
		 YearToSaveseatToRowSheetName=	'seatToRow'	+ 	(yearToCreate-1).toString();							
	       nm=supportWB.SheetNames.length; 
         sheetNum=supportWB.SheetNames.indexOf(YearToSaveseatToRowSheetName); 
			   lastYrseatToRowSheet=supportWB.Sheets['seatToRow']; 
         if(sheetNum != -1){ rsltToRow= '900';}   // requested saving already done
				else {
				
				supportWB.SheetNames[nm]=YearToSaveseatToRowSheetName; 
								
								newWs = creatSheet(seatToRowNumOfCols ,  seatToRowNumOfRows);  
								supportWB.Sheets[YearToSaveseatToRowSheetName]=newWs;
						
								YrSheet=supportWB.Sheets[YearToSaveseatToRowSheetName]; 
								
								for (rw=1; rw<seatToRowNumOfRows;rw++) {      // copy all rows to file to Save a 
								       rww=rw.toString();
								       Object.keys(seatToRowAmudot).forEach(function(key)  {   // copy all colomns

                            ptr1= seatToRowAmudot[key]+rww;
														vlu=' ';
														if( lastYrseatToRowSheet[ptr1]) vlu=lastYrseatToRowSheet[ptr1].v;
									        
														YrSheet[ptr1]={t:"s",v:vlu};
														YrSheet[ptr1].v=(YrSheet[ptr1].v).toString();  //make sure it is of str attribute
												
                         });	
								 
									}	// for	
									
							} //else		
														
							xlsx.writeFile(supportWB, supportTblsFilename);	
							
							
					
		return rsltWmn+'$'+rsltToRow;
									
	}
	



//----------------------------------------------------------
	function creatSheet( cols, rows) {
	var ws = {};
	var range = {s: {c:0, r:0}, e: {c:cols, r:rows }};
	for(var R = 0; R < rows; ++R) {
		for(var C = 0; C <cols; ++C) {
			
			var cell = {v: '   ' };
			
			var cell_ref = xlsx.utils.encode_cell({c:C,r:R});
			
			 cell.t = 's';
			
			ws[cell_ref] = cell;
		}
	}
	 ws['!ref'] = xlsx.utils.encode_range(range);
	return ws;
}

//----------------------------------------------------------	  
app.get('/getShulConfigPerm', function(req, res) {
	res.header("Access-Control-Allow-Origin", "*");
	 res.setHeader('Content-Type', 'text/html');
    configWS=  supportWB.Sheets['shulConfigPerm'];
    respns=getShulConfig( configWS); 
		res.send(rspns);	
		})


  
//----------------------------------------------------------	  
		
	
	app.get('/getShulConfig', function(req, res) {
	res.header("Access-Control-Allow-Origin", "*");
	 res.setHeader('Content-Type', 'text/html');
    configWS=  shulConfigerationWS;
		 rspns=getShulConfig( configWS);  
				res.send(rspns);	
					
	
			 
			  })	
																	


		 
		
		
	

//----------------------------------------------------------
		
		
function getShulConfig(configWS){

 rspns='';
						i=firstConfigRow;
						Istr=(i).toString();
						
						while(configWS[amudotOfConfig.fromSeat+Istr].v != '$$$'){
						  fromSt=configWS[amudotOfConfig.fromSeat+Istr].v;
						 if( Number(fromSt) ){
						// console.log('amudotOfConfig.ulam+Istr='+amudotOfConfig.ulam+Istr+'   fromSt='+fromSt);
						    ulam=configWS[amudotOfConfig.ulam+Istr].v;
							 //if (ulam.substr(0,1) != 'n'){nashim=0;} else nashim=1;
							 nashim= ulam.substr(0,1);
							 itmp=ulam.indexOf(' ');
		           tmp=ulam.substr(itmp+1,1);
							 UlamOrMartef='m';
							 if (tmp != 'm'){if (tmp=='e'){UlamOrMartef='n'} else UlamOrMartef='r'};  // 'e' == ezrat nashim
							 slantedX=configWS[amudotOfConfig.X_forSlantedRow+Istr].v;
							 if ( isNaN(slantedX) )slantedX='';
							 slantedY=configWS[amudotOfConfig.Y_forSlantedRow+Istr].v;
							 if ( isNaN(slantedY) )slantedY='';
						  rspns=rspns+configWS[amudotOfConfig.fromSeat+Istr].v+'@'
							+configWS[amudotOfConfig.toSeat+Istr].v+'@'
							+configWS[amudotOfConfig.reltvRowQual+Istr].v+'@'
							+configWS[amudotOfConfig.open_badSeats+Istr].v+'@'
							+configWS[amudotOfConfig.ezor+Istr].v+'@'
							+slantedX+'@'
							+slantedY+'@'
							+UlamOrMartef+'@'
							+	nashim+  '$';  
							}; 
						  i++;
							Istr=(i).toString();  
							};
        rspns=rspns.substr(0,rspns.length-1);  
				return rspns;		



}

//----------------------------------------------------------		
		
	
	app.get('/gizbar', function(req, res) {
	//res.header("Access-Control-Allow-Origin", "*");
	 res.setHeader('Content-Type', 'text/html');
	 var dbg;
		
		dbg=searchDebugParam('disable');  
		if ( ( dbg != -1 ) || ( ! initDone) ){
		          res.send(cache_get('real_index'));
							return;
							}
	 
	 
            res.send(cache_get('gizbar') );
        })	
				
				//-----------------------------------------------------------
app.get('/printBaseHtml', function(req, res) {
	//res.header("Access-Control-Allow-Origin", "*");
	 res.setHeader('Content-Type', 'text/html');
	   res.send(cache_get('prtBase') );  
	  
        })			
									
//-----------------------------------------------------------
app.get('/prtMartef', function(req, res) {
	//res.header("Access-Control-Allow-Origin", "*");
	 res.setHeader('Content-Type', 'text/html');
	 inputString=decodeURI(req.originalUrl);  
	 if(inputString.split('&')[1] == mngmntPASSW){  res.send(cache_get('prtMartef') )   }
	  else res.send(cache_get('errPasswd') );
        })			
				
//--------------------------------------------------------------
app.get('/prtRashi', function(req, res) {
	//res.header("Access-Control-Allow-Origin", "*");
	 res.setHeader('Content-Type', 'text/html');
	 inputString=decodeURI(req.originalUrl);  
	 if(inputString.split('&')[1] == mngmntPASSW){  res.send(cache_get('prtRashi') )   }
	  else res.send(cache_get('errPasswd') );
            
        })			
				
//-------------------------------------------------------

app.get('/prtNashim', function(req, res) {
	//res.header("Access-Control-Allow-Origin", "*");
	 res.setHeader('Content-Type', 'text/html');
	 inputString=decodeURI(req.originalUrl);  
	 if(inputString.split('&')[1] == mngmntPASSW){  res.send(cache_get('prtNashim') )   }
	  else res.send(cache_get('errPasswd') );
	
        })			
//------------------------------------------------------------------------------	

app.get('/ckpswMGMT', function(req, res) {
	res.header("Access-Control-Allow-Origin", "*");
	 res.setHeader('Content-Type', 'text/html');
	 
	 inputString=req.originalUrl.substr(12);  
	 if(inputString==mngmntPASSW){rspns='+++';} else rspns='---';
            res.send(rspns );
        })	


//------------------------------------------------------------------------------	
	app.get('/sendBrsrRprtedErr', function(req, res) {
	res.header("Access-Control-Allow-Origin", "*");
	 res.setHeader('Content-Type', 'text/html');
	 
	    maill='kehilatarielseats@gmail.com';
	    subj='errFromBrowser';
			txt=decodeURI(req.originalUrl);
	    sendMail(maill,subj,txt);
            res.send('+++' );
        })				
//------------------------------------------------------------------------------	

app.get('/getYearBeforeCurrent', function(req, res) {
	res.header("Access-Control-Allow-Origin", "*");
	 res.setHeader('Content-Type', 'text/html');
//	 initFromFiles('');
	 
	 var i, str;
	 str='';
	 
	
	 res.send(str );
        })	






//------------------------------------------------------------------------------

app.get('/getOverAssignedList', function(req, res) {
	res.header("Access-Control-Allow-Origin", "*");
	 res.setHeader('Content-Type', 'text/html');
//	 initFromFiles('');
	 
	 var i, str,tmp,name;
	 str='ok';
	 for (i=firstSeatRow;i<lastSeatRow+1;i++){ 
    row=i.toString();
		cell=requestedSeatsWorksheet[amudot.name+row];
		if (! cell)continue;
		name=simplifyName(cell.v);      /*.split('*');
		if ( name[1]  && name[2] ){tmp=name[1]+' '+hebrewLetters.vav+name[2] } else tmp=name[1]+name[2];
    name=name[0]+' '+tmp;
		*/		      
	  if(
	        ( Number(requestedSeatsWorksheet[amudot.menRosh+row].v) < Number(requestedSeatsWorksheet[amudot.numberOfAssignedSeatsRoshMen+row].v) )
			||  ( Number(requestedSeatsWorksheet[amudot.womenRosh+row].v) < Number(requestedSeatsWorksheet[amudot.numberOfAssignedSeatsRoshWomen+row].v) )
			||  ( Number(requestedSeatsWorksheet[amudot.menKipur+row].v) < Number(requestedSeatsWorksheet[amudot.numberOfAssignedSeatsKipurMen+row].v) )
			||  ( Number(requestedSeatsWorksheet[amudot.womenKipur+row].v) < Number(requestedSeatsWorksheet[amudot.numberOfAssignedSeatsKipurWomen+row].v) )
			)str=str+'+'+name;
	}  // for		
	
	 res.send(str );
        })	


//------------------------------------------------------------------------------	

// ck if registration is closed


app.get('/isRegistrationClosed', function(req, res) {
	res.header("Access-Control-Allow-Origin", "*");
	 res.setHeader('Content-Type', 'text/html');
	 initFromFiles('');
	 
	 isWomanString=isWoman.join('+');
	 rspns='---$ $'+isWomanString;
	 ptr=amudot.registrationClosedDateNTime+'2';  
	 tmp=delLeadingBlnks(requestedSeatsWorksheet[ptr].v); 
	 if(tmp)rspns='+++$'+tmp+'$'+isWomanString;
	 res.send(rspns );
        })	
				
	//------------------------------------------------------------------------------						
	
	app.get('/getPermanentSeatsList', function(req, res) {
	res.header("Access-Control-Allow-Origin", "*");
	 res.setHeader('Content-Type', 'text/html');
	 var membershipLevel,tmp;
	var tempList=[];
	 inputString=decodeURI(req.originalUrl).split('?')[1];  
	 inputString=inputString.split('$');
	 if (inputString[0] != moshavimPASSW){console.log('wrong password'); res.send('---' );return}
	 
	   k=0;
		 minMembershipLevel=Number(inputString[1]);
	   for (member=firstSeatRow; member<lastSeatRow+1; member++){ 
		    row=member.toString(); 
		   
		    nam=delLeadingBlnks(requestedSeatsWorksheet[amudot.name+row].v);
		    if (nam ){
				      tmp=requestedSeatsWorksheet[amudot.memberShipStatus+row].v;
							tmp=tmp.toString(); //make sure it is a string before using delLeadingBlnks
							
				      membershipLevel=Number(delLeadingBlnks(tmp)); 
							if(membershipLevel < minMembershipLevel)continue;    //does not deserve permanent seat
							
							
							tempList[k]=nam+'<>'+minimumName[member-firstSeatRow];
							prm=requestedSeatsWorksheet[amudot.permanentSeats+row];
							if(prm){
							prm=delLeadingBlnks(prm.v);
							if (prm  )tempList[k]=tempList[k]+'+'+prm;
							}  //if prm
							}   // if nam
				k++;
				}  // for
		rspns='+++'+tempList.join('$');
	  res.send(rspns );						
							
							
									
								
								 
 
	 })
	 
//------------------------------------------------------------------------------						
	
	app.get('/setPermanentSeatsList', function(req, res) {
	res.header("Access-Control-Allow-Origin", "*");
	 res.setHeader('Content-Type', 'text/html');
	var tempList=[];
	 inputString=decodeURI(req.originalUrl).split('?')[1];  
	 msgParts=inputString.split('$');
	 if (msgParts[0] != moshavimPASSW){console.log('wrong password'); res.send('---' )}
	 else {
	      for (i=1; i<msgParts.length;i++){
				    entry=msgParts[i].split('@');  
				    nam=entry[0];  
						rowNum=knownName(nam)[0]; if(rowNum==-1){console.log('in setPermanentSeatsList, not found name='+nam+'/'); continue;} 
						ptr=amudot.permanentSeats+rowNum.toString(); 
						requestedSeatsWorksheet[ptr]={t:"s",v:entry[1]};
				//		requestedSeatsWorksheet[ptr].v=entry[1]; 
			};			
	   
		xlsx.writeFile(workbook, XLSXfilename);  // update file
	  res.send('+++' );						
							
							
									
			}  // else						
								 
 
	 })
	 

//------------------------------------------------------------------------------	ckpswMoshavim				
				

app.get('/ckpswGIZBAR', function(req, res) {
	res.header("Access-Control-Allow-Origin", "*");
	 res.setHeader('Content-Type', 'text/html');
	 
	 inputString=req.originalUrl.substr(14);
	 if(inputString==gizbarPASSW){rspns='+++';} else rspns='---';
            res.send(rspns );
        })	

//------------------------------------------------------------------------------					
				

   app.get('/ckpswMoshavim', function(req, res) {
	res.header("Access-Control-Allow-Origin", "*");
	 res.setHeader('Content-Type', 'text/html');
	 
	 inputString=decodeURI(req.originalUrl).split('?')[1]; 
	
	 if(inputString==moshavimPASSW){rspns='+++';} else rspns='---';
            res.send(rspns );
        })	
//------------------------------------------------------------------------------		

var amudot_memberPersonalInfo=[amudot.name,  amudot.email,  amudot.addr,  amudot.phone, amudot.memberShipStatus];

 app.get('/manageMemberInfo_getAll', function(req, res) {
	res.header("Access-Control-Allow-Origin", "*");
	 res.setHeader('Content-Type', 'text/html');
	

	 	 initFromFiles('');
		 
	   rspns='+++';
	   for(i=firstSeatRow;i<lastSeatRow;i++){ 
         row=i.toString();
				 namePtr=amudot.name+row;
				 if (  (requestedSeatsWorksheet[namePtr])  && (delLeadingBlnks(requestedSeatsWorksheet[namePtr].v)) ){
				 rspns=rspns+'>';
				   for (j=0;j<	 amudot_memberPersonalInfo.length;j++){
					
					      ptr1=amudot_memberPersonalInfo[j]+row;  
							  rspns=rspns+delLeadingBlnks(requestedSeatsWorksheet[ptr1].v)+'$' ;
							} // for j	
					 rspns=rspns.substr(0,rspns.length-1);  // delete last $
				}  //if
			} // for i		 
        //   console.log('5180 rspns='+rspns);
					  res.send(rspns );
 })				
	
//------------------------------------------------------------------------------					
		 app.get('/isKnownName', function(req, res) {
	res.header("Access-Control-Allow-Origin", "*");
	 res.setHeader('Content-Type', 'text/html');	
	 var rNmA;
	 initFromFiles('');
	 nameToCheck=decodeURI(req.originalUrl).split('?')[1]; 
	 rNmA=isThisNameKnown	(nameToCheck);
	 
	 res.send(rNmA[0].toString()+'@'+rNmA[1]);
		
		
	}); 
//---------------------------------
function  	isThisNameKnown(nameToCheck){	
	var rNmA = new Array(); 
	var rNm,rn;
	var tempFamName,tmp,idx,familyNames;
	var strParts;        
	var strOriginalLength,indices, numberOfPops,tempFamName,firstIdx,i,j,nextIdx,confirmedIndices;
	var firstNamesArray, nameA, nameB,bothNames,rawList,cond1stWay,con2ndWay;
	
	 
	
	 // console.log('nameToCheck='+nameToCheck);
	
  rawList=[];  
	familyNames=[];
	bothNames=[];
	idx=0;
	for (i=firstSeatRow;i<lastSeatRow+1;i++) {
	   cell= requestedSeatsWorksheet[amudot.name+i.toString()];
		 if ( ! cell)  continue;  
		  if ( ! delLeadingBlnks(cell.v) )  continue;    
	    rawList[idx]=delLeadingBlnks(cell.v);
	    tmp=( rawList[idx]).split('*');
      familyNames[idx]=tmp[0];
			bothNames[idx]=[tmp[1],tmp[2]];
			idx++;
	//	console.log(	familyNames[idx-1]);
	}   // for i 
	
	nameParts=nameToCheck.split(' ');  
	strOriginalLength=nameParts.length;
	
	indices=[];
	numberOfPops=-1;
	while (nameParts.length){
	    numberOfPops++;
    	tempFamName=nameParts.join(' '); 
		  firstIdx=findIdxOfName(	familyNames,tempFamName,0);   // console.log('firstIdx='+firstIdx+' tempFamName='+tempFamName);
			if (firstIdx  == -1){ // this combination not found
			     nameParts.pop();  // remove the last part of the name in case it is a first name
					 continue;  // try a shorter name
					 }  // if
			// family name found. now look for all possible families with the same family name
			 // start looking for this name from the first family name
			 indices.push([firstIdx,numberOfPops]);
			 // console.log(' indices 1 ='+indices);
			
			 nextIdx=findIdxOfName(	familyNames,tempFamName,firstIdx+1);
			 while (nextIdx  != -1 ){
			   indices.push([nextIdx,numberOfPops]);  //console.log(' indices 2 ='+indices);
				
					nextIdx=	findIdxOfName(	familyNames,tempFamName,nextIdx+1)
					}  // while nextIdx
					
			 nameParts.pop();  // remove the last part of the name in case it is a first name		
			 
			 }   // while strparts.length
		//for(i=0;i<indices.length;i++)console.log('i='+i+' indices[i]='+indices[i]+' familyNames[indices[i][0]]='+familyNames[indices[i][0]]+' bothNames[indices[i][0]]='+bothNames[indices[i][0]]);
	
			 // now we have all indices of possible last name
			 confirmedIndices=[];
			 //  next step == prone all indices that have different first name(s)
			  for (i=0; i<indices.length; i++){
				     nextIdx=indices[i][0];
						 numberOfPops=indices[i][1];  // length (in tokens) of firstnames
						 nameParts=nameToCheck.split(' ');
						 firstNamesArray=nameParts.splice(strOriginalLength-numberOfPops,numberOfPops);
						 
						
						 
						 if ( ! firstNamesArray.length){ if(confirmedIndices.indexOf(nextIdx) == -1)confirmedIndices.push(nextIdx); continue};
						 // try all variations of how to generate two (or one) first names
						 for (j=0; j<firstNamesArray.length;j++){
						     nameA=firstNamesArray.splice(0,j).join(' ');
								 nameB=firstNamesArray.join(' '); 
							    if (nameB.substr(0,1) == hebrewLetters.vav) nameB=nameB.substr(1); 
									
									
									// try one way 
									cond1stWay= tryAMatch(nameA,nameB,bothNames[nextIdx]);
              //    console.log('nameA='+nameA+' nameB='+nameB+' cond1stWay='+cond1stWay);
								
									// try 2nd way
									
									   // change order
									con2ndWay=tryAMatch(nameB,nameA,bothNames[nextIdx]);
								//	console.log('nameA='+nameA+' nameB='+nameB+' con2ndWay='+con2ndWay);
									
									if ( cond1stWay ||  con2ndWay ) {		 if(confirmedIndices.indexOf(nextIdx) == -1)confirmedIndices.push(nextIdx); break;};
									nameParts=nameToCheck.split(' ');   // restore firstNamesArray
						      firstNamesArray=nameParts.splice(strOriginalLength-numberOfPops,numberOfPops);
							} // for j
					} // for i
											
				  
			if ( confirmedIndices.length != 1) { rNmA[0] = -1}else rNmA[0]=confirmedIndices[0];
			rNmA[1]=''; 
			for (i=0; i<confirmedIndices.length;i++)rNmA[1]=rNmA[1]+'$'+simplifyName(rawList[confirmedIndices[i]]);
			if (rNmA[1] )rNmA[1]=rNmA[1].substr(1);  
				
				
				return (rNmA);
				
		}		
			
	 
//-----------------------

function findIdxOfName(	familyNames,tempFamName,idxFrom){
var tmp,i,lngth,shortFamName;
lngth=tempFamName.length;

for (i=idxFrom;i< familyNames.length; i++){ 
  if(  familyNames[i].substr(0,lngth) ==  tempFamName)return i;
	}
	return -1;
	}
	
//-------------------
function tryAMatch(nameA,nameB,bothNames){
var shortBothNames,nameAL,nameBL,cond,cond1,cond2;

shortBothNames=[];

	                nameAL=nameA.length;
									nameBL=nameB.length;
									shortBothNames[0]=bothNames[0].substr(0,nameAL);
									shortBothNames[1]=bothNames[1].substr(0,nameBL);  
									
									cond1= ( (nameA) && (nameA != shortBothNames[0])  );  // if specified it has to be equal to the one in the DB
									cond2= ( (nameB) && (nameB != shortBothNames[1]) ); 
									
									cond=	(! cond1) && ( ! cond2);  
									return cond;
}									
  




		
	  
			
			
//------------------------------------------------------------------------------		
 app.get('/manageMemberInfo_update', function(req, res) {
	res.header("Access-Control-Allow-Origin", "*");
	 res.setHeader('Content-Type', 'text/html');
var 	inputString,tmp,i,j,updateRequest,roww,tmp;
	 	
		 inputString=decodeURI(req.originalUrl).split('?')[1];  
	
		 updateRequest=inputString.split('<>');;
		 
		  initFromFiles('');
			
			// update requests
			
			
			 ;
					row=knownName(updateRequest[0])[0];   // name before update
					if (row == -1){res.send('--- ' + updateRequest[0]+ ' is unknown');  return;};
					tmp=updateRequest[1].split('||')[1];
					if (  isThisNameKnown(simplifyName(tmp))[1] ){res.send('999 ' + updateRequest[1]+ ' is not unique');  return;};
					roww=row.toString();
					 for (j=1;j<	updateRequest.length;j++){
					 //requestedSeatsWorksheet[	 amudot_memberPersonalInfo[j]+row].v=updateRequest[j+1];
			    entryParts=updateRequest[j].split('||');
				   switch 	(entryParts[0]) {  
				     				 
				     case 'name': 
						          requestedSeatsWorksheet[amudot.name+row ].v=entryParts[1]; 
						          break;
							case 'email':
							        requestedSeatsWorksheet[amudot.email+row ].v=entryParts[1]; 
							        break;
											
							case 'phone':
							        requestedSeatsWorksheet[amudot.phone+row ].v=entryParts[1]; 
							        break;
											
							case 'addr':
							        requestedSeatsWorksheet[amudot.addr+row ].v=entryParts[1]; 
							        break;
											
							case 'membership':
							        requestedSeatsWorksheet[amudot.memberShipStatus+row ].v=entryParts[1]; 
							     break;
									 
						}   // switch			 																
						} // for j
		 
				 
	 sortDB();
	   
	 xlsx.writeFile(workbook, XLSXfilename);
	 
	
	
	  initFromFiles('');
		
	  res.send('+++' );
 })				
				
				
				
				
				
				
//------------------------------------------------------------------------------						
	app.get('/', function(req, res) {  
	//res.header("Access-Control-Allow-Origin", "*");
	 res.setHeader('Content-Type', 'text/html');
    
		var dbg;
		
		dbg=searchDebugParam('disable');  
	
		if ( ( dbg != -1 ) || ( ! initDone) ){
		          res.send(cache_get('real_index'));
							return;
							}
		   res.send(cache_get('index.html'));
        })			

//------------------------------------------------------------------------------	
app.get('/keepAlive', function(req, res) {
	//res.header("Access-Control-Allow-Origin", "*");
	 res.setHeader('Content-Type', 'text/html');
            res.send('ok' );
        })	
				
//------------------------------------------------------------------------------	
app.get('/moshavim', function(req, res) {
	//res.header("Access-Control-Allow-Origin", "*");
	 res.setHeader('Content-Type', 'text/html');
            res.send(cache_get('moshavim'));
        })					
//------------------------------------------------------------------------------							
				
	app.get('/iiiik', function(req, res) {
	//res.header("Access-Control-Allow-Origin", "*");
	 res.setHeader('Content-Type', 'text/html');
            res.send(cache_get('real_index') );
        })				
				start();
				
				
				
				
	