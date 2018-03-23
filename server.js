#!/bin/env node

var express = require('express');
var fs      = require('fs');
var http = require('http');

var path = require('path');
var xlsx = require('xlsx');

var nodemailer = require('nodemailer');


    /*  ================================================================  */

    /*
     *  Set up server IP address and port # using env variables/defaults.
     */
  function setupVariables  () {
        //  Set the environment variables we need.
       ipaddress = process.env.OPENSHIFT_NODEJS_IP;
        port      = process.env.OPENSHIFT_NODEJS_PORT || 8080;

        if (typeof ipaddress === "undefined") {
            //  Log errors on OpenShift but continue w/ 127.0.0.1 - this
            //  allows us to run/test the app locally.
            console.warn('No OPENSHIFT_NODEJS_IP var, using 127.0.0.1');
          ipaddress = "127.0.0.1";
        };
				
				console.log(ipaddress+'   '+port);
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
				
				zcache['prtMartef'] = fs.readFileSync('./martefBaseHtmlToPrint.html');  
				zcache['prtRashi'] = fs.readFileSync('./rashiBaseHtmlToPrint.html');  
				zcache['prtNashim'] = fs.readFileSync('./nashimBaseHtmlToPrint.html'); 
			  zcache['errPassw'] = fs.readFileSync('./errPassw.html');   
				zcache['gizbar'] = fs.readFileSync('./gizbar.html');
				zcache['okmsg'] = fs.readFileSync('./okmsg.html');
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
	function reportInputProblem(code){ reportErr(code+ txtCodes[1]);  }   //' נסה שנית. בעיה בנתונים'


//--------------------------------------------------------------------------------	
	function knownName(str){  
	var rNmA = new Array();           
	startingRow=3;
	//nameParts=str.split(' '); firstNameInRequest=(nameParts.length==2);   
	rNm=-1; 
	firstNamesString='';
	for (rn=0; rn<familyNames.length; rn++){
	     fnm_firstPossibility=delLeadingBlnks(familyNames[rn]);
	     fnm_parts=fnm_firstPossibility.split(' ');
	     
			 fnm_secondPossibilty=fnm_firstPossibility;
			 if (firstName[rn]){
			                     fnm_parts.pop();
													 fnm_secondPossibilty=fnm_parts.join(' ');
													 }
											
	   if (str == fnm_firstPossibility) { rNm=rn; break };
		 if (str == fnm_secondPossibilty) { firstNamesString=firstNamesString+	'$' +firstName[rn];	 };
	
			};
			
	if(rNm!=-1)rNm=rNm+startingRow;
	rNmA[0]=rNm;   rNmA[1]=firstNamesString;
	return rNmA;
	};
//--------------------------------------------------------------------------------
	
 function handleInput(inPairs){  
 
 seatRelevantChangeRequest=false;
 // name
 
  tmpName=inPairs[0].split('=');  name=tmpName[1];
  if (tmpName[0]== 'familyname'){
	 rowNum=knownName(name)[0];  
	   
		 if (rowNum ==-1 ){ reportInputProblem('000'); return false;} // 'שם לא מוכר'
		 }   
	   else { reportInputProblem('020'); return false};
			
	// email
	
		roww=rowNum.toString();

	tmpMail=	inPairs[1].split('=');
	 if (tmpMail[0]=='reqemail') {  
		email=tmpMail[1]; 
	  if (email){ ptr=amudot.email+roww;    requestedSeatsWorksheet[ptr].v=email;}
		} else {	reportInputProblem('002'); return false;};
	
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
		seats=seats.sort(sortOrder);
		lastSeatsRequest=(requestedSeatsWorksheet[ptr].v).split('+');
		changeInRequest=false;
		if(seats.length != lastSeatsRequest.length){
		                changeInRequest=true; }
					else {
					  for  (i=0; i<seats.length; i++)if (seats[i].split('_')[0] != 	lastSeatsRequest[i].split('_')[0])changeInRequest=true; 
						}
			
			if(	changeInRequest ){
			  						
										seatRelevantChangeRequest=true;
										}
		 requestedSeatsWorksheet[ptr].v=seats.join('+');
		
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
	
	
	
	update_namesForSeat(roww);
	
	   
		
// write detailed request	
	xlsx.writeFile(workbook, XLSXfilename);
	
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
	

  function setSeatOccupationLevel(holiday){     // holiday == 0 => both; holiday ==1 => rosh' holiday ==2 => kipur
	  
	  seatOcuupationLevel.forEach(setToZero);; // clear previous values
	 for (ii=0; ii<lastSeatNumber+1;ii++)if (alreadyAssignedSeatsRosh[ii] || alreadyAssignedSeatsKipur[ii]){
	                 combinedAlreadyAssigned[ii]=true} else combinedAlreadyAssigned[ii]=false;
									 
		 for (member=firstSeatRow; member<lastSeatRow+1; member++){ 
		    sMember=member.toString();  
				
				 toAssgnRoshMen=Number(requestedSeatsWorksheet[amudot.menRosh+sMember].v)-Number(requestedSeatsWorksheet[amudot.numberOfAssignedSeatsRoshMen+sMember].v);
 				 toAssgnRoshWomen=Number(requestedSeatsWorksheet[amudot.womenRosh+sMember].v)-Number(requestedSeatsWorksheet[amudot.numberOfAssignedSeatsRoshWomen+sMember].v);
 				 toAssgnKipurMen=Number(requestedSeatsWorksheet[amudot.menKipur+sMember].v)-Number(requestedSeatsWorksheet[amudot.numberOfAssignedSeatsKipurMen+sMember].v);
 				 toAssgnKipurWomen=Number(requestedSeatsWorksheet[amudot.womenKipur+sMember].v)-Number(requestedSeatsWorksheet[amudot.numberOfAssignedSeatsKipurWomen+sMember].v);

		
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
                        Date(Date.now() ), ipaddress, port);
        });
    };

//  ----- handle member info request ------------------------

function memberInfo(requestor,inputString){
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
		 
		             ocupAdd=ocupAdd+'-'+namesForSeat[i];  //  get all names for seat to client
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
	
		 if (rowNum ==-1 ){  respns='---000'+rawName[1];}    // 'שם לא מוכר'
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
						 if( ! inputForMemberExists) respnArray[positionInMsg.minyanMuadafNashim]='1';
						 respnArray[positionInMsg.esberNashim]=requestedSeatsWorksheet[amudot.preferedExplanationW+Row].v;
						 respnArray[positionInMsg.minyanMuadafGvarim]=(requestedSeatsWorksheet[amudot.preferedMinyanM+Row].v).substr(0,1);
						 if( ! inputForMemberExists) respnArray[positionInMsg.minyanMuadafGvarim]='1';
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
var reqMinyan = new Array;
var txtCodes = new Array;
var seatToRow = new Array;
var isWoman =new Array;
var startingRow;
var errorNumber;
var seats = new Array;
var incompatibilty;
var hasStillToPay;
var inputString;
var namesForSeat = new Array;
closeRegistrationDate =new Date;
var afterClosingDate;
var alreadyAssignedSeatsRosh = new Array;
var alreadyAssignedSeatsKipur = new Array;
var combinedAlreadyAssigned= new Array;

var debugIsOn = false;
var debugparam=''; 

var startingRowstsfction=6 ;
var stsfctionFamilyNames = new Array;

var  moedCode
var	 SidurUlam
var	 shlavBFilter;
var	 notPaidFilter;
var	 doneWithFilter;
/*
var maxCountRashiMen=0; 
var maxCountRashiWmn=0;
var maxCountMartefMen=0; 
var maxCountMartefWmn=0;
	*/
	
var 	maxCountSeats=[[0,0],[0,0]]; // rashi-martef  /  gvarim-nashim
var firstName = new Array;


var assignedPerUlam=[[[0,0],[0,0]],[[0,0],[0,0]]];  //[ulam][moed][gvarim-nashim]

var amudot ={name:'A',registrationClosedDateNTime:'C',requestDate:'D',email:'G',addr:'H',phone:'I',
              menRosh:'J',menKipur:'K',womenRosh:'L',womenKipur:'M',preferedMinyanW:'N',
              preferedExplanationW:'O',preferedMinyanM:'P',preferedExplanationM:'Q',cmnts:'R',
							markedSeats:'S',numberMarkedMen:'T',numberMarkedWomen:'U',notAssignedMarkedSeatsRosh:'V',
							notAssignedMarkedSeatsKipur:'W',NumberOfNotAssignedMarkedSeatsMen:'X', NumberOfNotAssignedMarkedSeatsWomen:'Y',
							assignedSeatsRosh:'Z',assignedSeatsKipur:'AA',numberOfAssignedSeatsRoshMen:'AB',numberOfAssignedSeatsRoshWomen:'AC',
							numberOfAssignedSeatsKipurMen:'AD',numberOfAssignedSeatsKipurWomen:'AE',tashlum:'AF',tashlumPaid:'AG',
							stsfctnInFlrLastYrWmn:'AI',stsfctnInFlrLastYrMen:'AJ',lstYrSeat:'AK',
							stsfctnInFlr2YRSAgoYrWmn:'AL',stsfctnInFlr2YRSAgoYrMen:'AM',TwoYRSAgoSeat:'AN',stsfctnInFlr3YRSAgoYrWmn:'AO',
							stsfctnInFlr3YRSAgoYrMen:'AP',ThreeYRSAgoSeat:'AQ',
							issueInFloorWmn:'AR',  issueinFloorMen:'AS',  issueBetweenFloors:'AT',   
							memberShipStatus:'AU',nashimMuadaf:'AV',gvarimMuadaf:'AW',stsfctnInFlrThisYrWmn:'AX',
							stsfctnInFlr3ThisYrMen:'AY',ThisYRSSeat:'AZ'
							};
	// 2 lists of colomns both for membersrequests and stsfction. first is membersrequests						
var amudotForStsfctionUpload={	
              stsfctnInFlr2YRSAgoYrWmn:['AL','B'],stsfctnInFlr2YRSAgoYrMen:['AM','C'],TwoYRSAgoSeat:['AN','D'],stsfctnInFlr3YRSAgoYrWmn:['AO','E'],
              stsfctnInFlr3YRSAgoYrMen:['AP','F'],ThreeYRSAgoSeat:['AQ','G'],memberShipStatus:['AU','N']
							};
							
							
amudotForStsfctionDownload= { name:['A','A'],menRosh:['G','O'],menKipur:['H','P'],womenRosh:['I','Q'],womenKipur:['J','R'], 
              stsfctnInFlrLastYrWmn:['AI','B'], stsfctnInFlrLastYrMen:['AJ','C'],LastYrSeat:['AK','D'],                        
              stsfctnInFlr2YRSAgoYrWmn:['AL','E'],stsfctnInFlr2YRSAgoYrMenn:['AM','F'],TwoYRSAgoSeat:['AN','G'],stsfctnInFlr3YRSAgoYrWmn:['AO','H'],
              stsfctnInFlr3YRSAgoYrMen:['AP','I'],ThreeYRSAgoSeat:['AQ','J'],memberShipStatus:['AU','N']
							};
                             


		
var	positionInMsg={email:0,	addr:1,phone:2,	gvarimRoshHashana:3,gvarimKipur:4,nashimRoshHashana:5,nashimKipur:6,
		                minyanMuadafNashim:7, esberNashim:8,minyanMuadafGvarim:9,esberGvarim:10,moreComments:11,
									 requestedSeats:12,requestDate:13,assignedSeatsRosh:14,assignedSeatsKipur:15,tashlum:16,tashlumPaid:17,
									 stsfctnInFlr2YRSAgoYrWmn:18,stsfctnInFlr2YRSAgoYrMen:19,TwoYRSAgoSeat:20,stsfctnInFlr3YRSAgoYrWmn:21,
                   stsfctnInFlr3YRSAgoYrMen:22,ThreeYRSAgoSeat:23,memberShipStatus:24,stsfctnInFlrLastYrWmn:25,stsfctnInFlrLastYrMen:26,
									 LastYrSeat:27,issueInFloorWmn:28,  issueinFloorMen:29,  issueBetweenFloors:30,nashimMuadaf:31,gvarimMuadaf:32
									 };
											  
	
var sortWeightsPtr={vetek:'F1',personalIssue:'F2',satisfactionHistory:'F3',satisfactionInFloor:'F4',horizontalDistance:'F5',
                    lastYearVS2YearsAgo:'F6',numberOfRequestedSeats:'F7',requestedSeatsPerFamilySize:'F8',Baby:'F10'}	
											

app=express();
initialize();

XLSXfilename=	process.env.OPENSHIFT_DATA_DIR+'membersRequests.xlsx';
EmptyXLSXfilename=	process.env.OPENSHIFT_DATA_DIR+'EmptymembersRequests.xlsx';           
seatsOrderedFileName=	process.env.OPENSHIFT_DATA_DIR+'seatsOrdered.xlsx';
errPasswFilename=process.env.OPENSHIFT_DATA_DIR+'empty.xlsx';
SortingDatafilename=	process.env.OPENSHIFT_DATA_DIR+'SortingData.xlsx';  

BackupFilename= process.env.OPENSHIFT_DATA_DIR+'BackupMembersRequests.xlsx';       

//-----------------------init gmail ------------------------------------------
   

    var transporter = nodemailer.createTransport({
        service: 'Gmail',
        auth: {
            user: 'kehilatarielseats@gmail.com', // Your email id
            pass: 'kehila11' // Your password
        }
    });
		
//--------------------------------------------------------------------------  

//read error codes and Seat to Row from supportTables.xlsx            


var supportWB=xlsx.readFile('supportTables.xlsx');
var errCodeWS=supportWB.Sheets['errorCodes'];
for (i=1; i<50; i++){
 ptr1='A'+(i).toString();
 if (errCodeWS[ptr1].v == '$$$') break;
 ptr2='B'+(i).toString();
 txtCodes[i]=errCodeWS[ptr2].v;    
      };
			
	lastSeatNumber=0;
var seatToRowWS=supportWB.Sheets['seatToRow'];
var isWomanWS=supportWB.Sheets['IsWoman'];
for (i=1; i<1500; i++){
 ptr1='A'+(i).toString();
 if (seatToRowWS[ptr1].v == '$$$'){lastSeatNumber=i-1; break; }
 alreadyAssignedSeatsRosh[i]=' '; 
 alreadyAssignedSeatsKipur[i]=' ';
 seatToRow[i]=seatToRowWS[ptr1].v;
 if (seatToRow[i] != 'NAN'){ 
    ptr1='A'+(seatToRow[i]).toString();
 isWoman[i]=isWomanWS[ptr1].v;
       } 
      };
			
	var nashimGvarimRegions=[];
	

	// get seats regions in martef and main; gvarim nashim; and count number of seats. gvarim nashim, in each region
	row=3;
	roww=row.toString();
	k=-1;
	while (isWomanWS['D'+roww].v != '$$$'){
	  k++; if(k>20){console.log('err in eizorei gvarim nashim / no $$$ as end of input'); break};
		nashimGvarimRegions[k]=new Array(5);
	  nashimGvarimRegions[k][0]=isWomanWS['C'+roww].v;   //ulam
		nashimGvarimRegions[k][1]=Number(isWomanWS['D'+roww].v);  //from seat - nashim
		nashimGvarimRegions[k][2]=Number(isWomanWS['E'+roww].v);  //to seat - nashim
		if(nashimGvarimRegions[k][1]!=nashimGvarimRegions[k][2]){    //count number of seats - nashim
		   tempCount=0;
			 for (j=nashimGvarimRegions[k][1]; j<nashimGvarimRegions[k][2]+1;j++)if (seatToRow[j] != 'NAN')tempCount++;
			 if (nashimGvarimRegions[k][0] == 'main'){maxCountSeats[0][0]=maxCountSeats[0][0]+tempCount} else maxCountSeats[1][0]=maxCountSeats[1][0]+tempCount;
			 }
		nashimGvarimRegions[k][3]=Number(isWomanWS['F'+roww].v);   //from seat - gvarim
		nashimGvarimRegions[k][4]=Number(isWomanWS['G'+roww].v);    //to seat - gvarim
		if(nashimGvarimRegions[k][3]!=nashimGvarimRegions[k][4]){    //count number of seats - gvarim
		   tempCount=0;
			 for (j=nashimGvarimRegions[k][3]; j<nashimGvarimRegions[k][4]+1;j++)if (seatToRow[j] != 'NAN')tempCount++;
			 if (nashimGvarimRegions[k][0] == 'main'){maxCountSeats[0][1]=maxCountSeats[0][1]+tempCount} else maxCountSeats[1][1]=maxCountSeats[1][1]+tempCount;
			 }
		
		row++;
	  roww=row.toString();
		}
	//console.log ('maxCountSeats='+maxCountSeats);

	//
		

	var passwordsWS=supportWB.Sheets['passwords'];
	mngmntPASSW=passwordsWS['B1'].v;
	gizbarPASSW=	passwordsWS['B2'].v;	



             
// -----------------------------------------------------------------------
		
			
//
// if first time make sure that an empty membersRequests exists
    stats = fs.statSync(XLSXfilename);  
   sts=stats.isFile();
    if (! sts){
  
		tmpfile=fs.readFileSync('EmptymembersRequests.xlsx');
	fs.writeFileSync(XLSXfilename, tmpfile);
	fs.writeFileSync(EmptyXLSXfilename, tmpfile);
	
	  console.log('membersRequests file was initialized');
}
  //
//
var seatOcuupationLevel = new Array;
for(i=1; i<lastSeatNumber+1; i++){
      seatOcuupationLevel[i]=0;    // clear and set array size 
			namesForSeat[i]='$/';
			};
			
			
var workbook = xlsx.readFile(XLSXfilename);
var emptymemberRequests=xlsx.readFile('EmptymembersRequests.xlsx');


	var requestedSeatsWorksheet = workbook.Sheets['HTMLRequests']; 
	var emptyRequestedSeatsWorksheet = emptymemberRequests.Sheets['HTMLRequests']; 
	// check if workbook is obsolete
	fileIsOK=true;
	Object.keys(amudot).forEach(function(key)  {

                            ptr1= amudot[key]+'1';
														if (requestedSeatsWorksheet[ptr1] && emptyRequestedSeatsWorksheet[ptr1] ){
									          if ( requestedSeatsWorksheet[ptr1].v != emptyRequestedSeatsWorksheet[ptr1].v )fileIsOK=false;
													                       }		
													
													 
                         });	


if ( !  fileIsOK ){
            tmpfile=fs.readFileSync('EmptymembersRequests.xlsx');
	          fs.writeFileSync(XLSXfilename, tmpfile);
	          fs.writeFileSync(EmptyXLSXfilename, tmpfile);
						var workbook = xlsx.readFile(XLSXfilename);
	          var requestedSeatsWorksheet = workbook.Sheets['HTMLRequests']; 
						console.log('membersRequests file was initialized');
}

initValuesOutOfHtmlRequestsXLSX_file();   //init values


CountAssignedPerMoed_PerUlam();

//open stsfction file and set stsfctionFamilyNames  

 var  stsfctionWB=xlsx.readFile('SortingData.xlsx');
 stsfctionSheets=stsfctionWB.SheetNames;
 var d = new Date();
 var crrntYyr = d.getFullYear();
 var lastYr=crrntYyr-1;
 var crrntYrSheetName='dataFor'+crrntYyr.toString();
 var lastYrSheetName='dataFor'+lastYr.toString();


if ( initSatisfactionFile() == '---')console.log('err in stsfction file list of names');
     
 var stsfctionWorkSheet = stsfctionWB.Sheets[lastYrSheetName]; 

	for (i=6;i<200;i++){ 
	 row=i.toString();
	 pointerCell=amudot.name+row;   //  name always in 'A'
 
	 	 cell=stsfctionWorkSheet[pointerCell]; 
	  if(! cell) continue;
		if ( ! delLeadingBlnks(cell.v) ) continue;
		if ( firstSeatRow == 0) startingRowstsfction=i;   // first name  row
	     stsfctionFamilyNames[i- startingRowstsfction]=cell.v; 
  }

	// get  sortWeights
sortWeightsSheet=stsfctionWB.Sheets['sortWeights'];   

vetekWeight=Number(sortWeightsSheet[sortWeightsPtr.vetek].v);
personalIssueWeight=Number(sortWeightsSheet[sortWeightsPtr.personalIssue].v);
satisfactionHistoryWeight=Number(sortWeightsSheet[sortWeightsPtr.satisfactionHistory].v);
satisfactionInFloorWeight=Number(sortWeightsSheet[sortWeightsPtr.satisfactionInFloor].v);
horizontalDistanceWeight=Number(sortWeightsSheet[sortWeightsPtr.horizontalDistance].v);
lastYearVS2YearsAgoWeight=Number(sortWeightsSheet[sortWeightsPtr.lastYearVS2YearsAgo].v);
numberOfRequestedSeatsWeight=Number(sortWeightsSheet[sortWeightsPtr.numberOfRequestedSeats].v);
requestedSeatsPerFamilySizeWeight=Number(sortWeightsSheet[sortWeightsPtr.requestedSeatsPerFamilySize].v);
BabyWeight=Number(sortWeightsSheet[sortWeightsPtr.Baby].v);


var lastBackupBefore=0;
backupEvery=24*6;  // every 24 hours
setTimeout(backupRequests, 600000);	//check every 10 minutes

//-------------------------------------------------------------
function backupRequests(){
  
   lastBackupBefore++;
	 if(lastBackupBefore > 0){// 
	 lastBackupBefore=0;
	 xlsx.writeFile(workbook, BackupFilename);
	 console.log('backup created');
	 
	 }
setTimeout(backupRequests, 600000);	//check every 10 minutes


}	 


//-------------------------------------------------------------
function initValuesOutOfHtmlRequestsXLSX_file(){
   firstName=[];
	 	 firstSeatRow=0;
	 for (i=2;i<200;i++){ 
	 row=i.toString();
	 pointerCell=amudot.name+row;

	 
	 cell=requestedSeatsWorksheet[pointerCell]; 
	  if(! cell) continue;
		cell.v=delLeadingBlnks(cell.v);
		if ( ! cell.v )continue;
		if ( firstSeatRow == 0) firstSeatRow=i;   // first name  row
		
	     famName=cell.v; 
			 if (famName == '$$$'){ lastSeatRow=i-1; break;}
			 famParts= famName.split(' ');
			 famName=famParts.join(' '); // remove un necessary blanks
			 if (famName.substr(famName.length-1,1)=='*') {
			     famName=famName.substr(0,famName.length-1);
					 theFirstName=famParts[famParts.length-1];
					 theFirstName=theFirstName.substr(0,theFirstName.length-1);
					 firstName[i-firstSeatRow]=theFirstName;
			     } 
					 else firstName[i-firstSeatRow]='';
					 
					 
	     familyNames[i-firstSeatRow]= famName; 
			  
		   mRosh=Number(requestedSeatsWorksheet[amudot.menRosh +row].v);
				 wRosh=Number(requestedSeatsWorksheet[amudot.womenRosh +row].v);
				 mKipur=Number(requestedSeatsWorksheet[amudot.menKipur +row].v);
				 wKipur=Number(requestedSeatsWorksheet[amudot.womenKipur +row].v); 
				
				 if ( (mRosh + wRosh + mKipur + wKipur) == 0) continue; // member did not input a request

				 closeSeats(1,i);
				closeSeats(2,i);  
						 
	   }
		
	 if (i>190)reportAnError('no $$$ at end of family names'); 
	 
	 for (i=1; i<lastSeatNumber+1; i++){
               alreadyAssignedSeatsRosh[i]=' '; 
               alreadyAssignedSeatsKipur[i]=' '; 
							 }
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


function update_namesForSeat(row){
debug1=false;     //if (row  =='173') debug1=true; 
  nm=delLeadingBlnks( requestedSeatsWorksheet[amudot.name +row].v);

 
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
 
 var counts=[];
 var listToSend=[];
 var tempList=[];
 idx=0;
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
		
		// prepare for future "doneWith" filtering
		
		   toAssgnRoshMen=Number(requestedSeatsWorksheet[amudot.menRosh+row].v)-Number(requestedSeatsWorksheet[amudot.numberOfAssignedSeatsRoshMen+row].v);
 			 toAssgnRoshWomen=Number(requestedSeatsWorksheet[amudot.womenRosh+row].v)-Number(requestedSeatsWorksheet[amudot.numberOfAssignedSeatsRoshWomen+row].v);
 			 toAssgnKipurMen=Number(requestedSeatsWorksheet[amudot.menKipur+row].v)-Number(requestedSeatsWorksheet[amudot.numberOfAssignedSeatsKipurMen+row].v);
 			 toAssgnKipurWomen=Number(requestedSeatsWorksheet[amudot.womenKipur+row].v)-Number(requestedSeatsWorksheet[amudot.numberOfAssignedSeatsKipurWomen+row].v);
		// name not filtered
		
		counts=counSeatsInEzor(row,moed,SidurUlam);
		nameToKeep= delLeadingBlnks(requestedSeatsWorksheet[amudot.name +row].v);
		if (nameToKeep.substr(nameToKeep.length-1,1) == '*') nameToKeep=nameToKeep.substr(0,nameToKeep.length-1);
		tempList[idx]=nameToKeep+'$'+calcSortParam(row)[0]+'$'+calcSortParam(row)[1]+
		                '$'+counts[0].toString()+'$'+counts[1].toString()
										+'$'+toAssgnRoshMen.toString()+'$'+toAssgnRoshWomen.toString()+'$'+toAssgnKipurMen.toString()+'$'+toAssgnKipurWomen.toString();
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
	
//console.log('tempList_2='+tempList_2);

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
		    if ( (doneWithFilter=='true') && doneWithFlag)continue;
				listToSend[idx]=vlu[0]; // strip not required info
				idx++; 
		}  //doneWithFilter
		

 strr=listToSend.join('$');   
 return strr;
 }
 

//--------------------------------------------------------------------------     
function calcSortParam(row){
 var calcResult=[];
 
 //   calc family sort calue
 
// part personal problem family between floors
part1= personalIssueWeight*Number(requestedSeatsWorksheet[amudot.issueBetweenFloors+row].v);

// part vetek
part2=vetekWeight*Number(requestedSeatsWorksheet[amudot.memberShipStatus+row].v);

//part mishkal koma in the past
MakomMinus3Yrs=Number(requestedSeatsWorksheet[amudot.ThreeYRSAgoSeat+row].v);
MakomMinus2Yrs=lastYearVS2YearsAgoWeight*Number(requestedSeatsWorksheet[amudot.TwoYRSAgoSeat+row].v);
MakomMinus1Yrs=lastYearVS2YearsAgoWeight*lastYearVS2YearsAgoWeight*Number(requestedSeatsWorksheet[amudot.lstYrSeat+row].v);

part3=satisfactionHistoryWeight*(MakomMinus3Yrs+MakomMinus2Yrs+MakomMinus1Yrs);

// part mispar mekomo mevukash
numOfRequestedSeats=Number(requestedSeatsWorksheet[amudot.menRosh+row].v)+Number(requestedSeatsWorksheet[amudot.menKipur+row].v)
                  +Number(requestedSeatsWorksheet[amudot.womenRosh+row].v)+Number(requestedSeatsWorksheet[amudot.womenKipur+row].v);
part4=numberOfRequestedSeatsWeight*numOfRequestedSeats;

// part mispar mekomo mevukash vs family size
part5=requestedSeatsPerFamilySizeWeight*numOfRequestedSeats;

//sum all for first sort value
calcResult[0]=part1+part2+part3-part4-part5+10000;



// calc nashim+gvarim issue in floor sort value

part6=personalIssueWeight*(Number(requestedSeatsWorksheet[amudot.issueInFloorWmn+row].v)+Number(requestedSeatsWorksheet[amudot.issueinFloorMen+row].v));

// part satisfaction history
stsfctnMinus3Yrs=Number(requestedSeatsWorksheet[amudot.stsfctnInFlr3YRSAgoYrWmn+row].v)+Number(requestedSeatsWorksheet[amudot.stsfctnInFlr3YRSAgoYrMen+row].v);
stsfctnMinus2Yrs=lastYearVS2YearsAgoWeight*(Number(requestedSeatsWorksheet[amudot.stsfctnInFlr2YRSAgoYrWmn+row].v)+Number(requestedSeatsWorksheet[amudot.stsfctnInFlr2YRSAgoYrMen+row].v));
stsfctnMinus1Yrs=lastYearVS2YearsAgoWeight*lastYearVS2YearsAgoWeight*(Number(requestedSeatsWorksheet[amudot.stsfctnInFlrLastYrWmn+row].v)
                       +Number(requestedSeatsWorksheet[amudot.stsfctnInFlrLastYrMen+row].v));

 part7=  stsfctnMinus3Yrs+  stsfctnMinus2Yrs+ stsfctnMinus1Yrs;
 
 // part baby
 
 if ( (requestedSeatsWorksheet[amudot.preferedMinyanW+row].v).substr(0,1) =='0'){
       SourceBirthDate=  (requestedSeatsWorksheet[amudot.preferedExplanationW+row].v);
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
 		 
return calcResult;


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
	var alreadyAssignedTemp = new Array;
	
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
	if (  delLeadingBlnks(strOfSeats)){ // return;
	tmpAssigned=strOfSeats.split('+');  
	for (ii=0; ii < tmpAssigned.length; ii++){
	 seatNm=Number(tmpAssigned[ii]);  
	 alreadyAssignedTemp[seatNm]=nameForSeat;
	};	
	
	tArray=[];
	tArray=countMenAndWomenAssignedSeats(row);
	
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
  //   console.log(assignedPerUlam);

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
		 eizorAndGender=getEizorForSeat(Cseat);
		 eizor=['main','martef'].indexOf(eizorAndGender[0]);
		 assignedPerUlam[eizor][moed][eizorAndGender[1]]++;
	}
	
}


//--------------------------------------------------------------------------    
function getEizorForSeat(seatNum){  
  for (ll=0;ll<nashimGvarimRegions.length;ll++){
	
	     if ( (seatNum >= nashimGvarimRegions[ll][1])  && (seatNum<= nashimGvarimRegions[ll][2]) )return [nashimGvarimRegions[ll][0],0];
    	 if ( (seatNum >= nashimGvarimRegions[ll][3])  && (seatNum<= nashimGvarimRegions[ll][4]) )return [nashimGvarimRegions[ll][0],1];
			 }
			 console.log('err in nashimGvarimRegions');

}

//--------------------------------------------------------------------------    
	function sendMsgToKehilatArielSeatsGmail(titl,Msg){
	
		 var mailOptions = {
    from: 'kehilatarielseats@gmail.com', // sender address
    to: 'kehilatarielseats@gmail.com', // list of receivers
    subject: titl, // Subject line
    text: Msg //, // plaintext body
                       };
									    
   transporter.sendMail(mailOptions, function(error, info){
    if(error)  console.log('send mail reported an error=='+error);
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
  passW=inputString.split('-')[1]; tt=inputString.split('-'); 
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

app.get('/seatsOrderedXLS', function(req, res) {
	res.header("Access-Control-Allow-Origin", "*");
	
	inputString=decodeURI(req.originalUrl); 
  passW=inputString.split('-')[1]; //tt=inputString.split('-'); 
	if (passW == mngmntPASSW){  fileToSendName= 'seatsOrdered.xlsx';  fileToRead=seatsOrderedFileName; generate_seatsOrderedXLS();}
				else {fileToSendName='empty.xlsx'; fileToRead='empty.xlsx'}
				
				
				
        res.setHeader('Content-type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
        res.setHeader('Content-Disposition', 'attachment; filename=' + fileToSendName);
	      var fileR = fs.readFileSync(fileToRead, 'binary');
        res.setHeader('Content-Length', fileR.length);
        res.write(fileR, 'binary');
        res.end();
      
 
});
//---------------------------------------------------------------------------------	
 //emptyHazmanatmekomotFileName=	process.env.OPENSHIFT_DATA_DIR+'hazmanatMekomotEmpty.xlsx';  
	 function generate_seatsOrderedXLS(){
	 
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
	 
	 var amudotHazmana=['A','B','C','D','E','F','G','H','I','J','K','L','M','N','O'];
	 nameslist=familyNames.sort();
	 for (ik=0; ik<nameslist.length;ik++){
	     memberDataName[ik]=nameslist[ik];
			 rowNum=knownName(memberDataName[ik])[0];
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
		
		
		// update report date	
		offset=0;
		var d= Date();
		var mnthLngth=[31,28,31,30,31,30,31,31,30,31,30,31];	
		var monthNames=['Jan','Feb','Mar','Apr','May','Jun','Jul','Aug','Sep','Oct','Nov','Dec'];	 
			var n = new Date; 	
			dParts=n.toString().split(' ');  
			localTimeZoneDiffToZero= Number(dParts[5].substr(3,3));
			offset=2-localTimeZoneDiffToZero; //  Israel is GMT+2
			HR=Number(dParts[4].substr(0,2))+offset;
			dy=Number(dParts[2]);
			mnth=monthNames.indexOf(dParts[1])+1;
			yr=Number(dParts[3]);
			if (yr/4 == Math.round(yr/4)){mnthLngth[1]=29} else mnthLngth[1]=28;
			
			if (HR > 23){   HR=HR-24;   Dy++};
			if(dy > mnthLngth[mnth-1] ) {dy=1; mnth++};
			if (mnth>12){mnth=1; yr++};
			newDate=dy.toString()+'/'+mnth.toString()+'/'+yr.toString();
			
			
			hazmanotSheet['B4'].v=newDate;
			
			
	 // write the data into a new file
	 
	 xlsx.writeFile(tmpfile, seatsOrderedFileName);
	 
	 }



//---------------------------------------------------------------------------------	  

app.get('/addMember', function(req, res) {
	res.header("Access-Control-Allow-Origin", "*");

	 inputString=decodeURI(req.originalUrl);  
	 inputPairs=inputString.split('&');
	 
 
   passW=inputPairs[4];
	 if (passW == mngmntPASSW){
	 lastNm=delLeadingBlnks(inputPairs[1]);
   frstNm= delLeadingBlnks(inputPairs[2]);
	 
	 // is there such a name already in the system?
	 if(frstNm){   nameToCheck=lastNm+' '+ frstNm; } else nameToCheck=lastNm;
	 reslt=knownName(nameToCheck);  
	 if (reslt[0] ==  -1) {   // either not found or there are others with this name
	   if ( reslt[1] ){   //   there are others with this name
		       //  demand specifying first name
					 res.send('--1');
	         return;
					 }
		} else // not -1; found such a name; request user to add first name to existing name
				           {
				   res.send('--2');
	         return;
					 }
			
			
					    	 
	 lastSeatRow++;
	 
	 if (frstNm){ 
	        newName=lastNm+' '+ frstNm+'*';
					familyNames[lastSeatRow-firstSeatRow]= lastNm+' '+ frstNm;
					firstName[lastSeatRow-firstSeatRow]=frstNm;
					}  else {
					      newName=lastNm;
								familyNames[lastSeatRow-firstSeatRow]= lastNm;
					      firstName[lastSeatRow-firstSeatRow]='';
                  }
	 roww=lastSeatRow.toString()
   pointerCell=amudot.name+roww; 
	 requestedSeatsWorksheet[pointerCell].v=newName; 	
	 ptr=amudot.memberShipStatus+roww; 
	 requestedSeatsWorksheet[ptr].v=inputPairs[3];
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
	 
   
	 xlsx.writeFile(workbook, XLSXfilename);
	 
	 var Emptyworkbook = xlsx.readFile(EmptyXLSXfilename); 
   var EmptyrequestedSeatsWorksheet = Emptyworkbook.Sheets['HTMLRequests']; 
	 EmptyrequestedSeatsWorksheet[pointerCell].v=newName; 	
   
	 xlsx.writeFile(workbook, EmptyXLSXfilename);
	 
	
	 
	 res.send('+++');
	 }
	else res.send('999' ); 
	 
});

//---------------------------------------------------------------------------------	  
app.get('/addFirstName', function(req, res) {
	res.header("Access-Control-Allow-Origin", "*");

	 inputString=decodeURI(req.originalUrl);  
	 inputPairs=inputString.split('&');
	 
 
   passW=inputPairs[3];
	 if (passW == mngmntPASSW){
	 lastNm=delLeadingBlnks(inputPairs[1]);
   frstNm= delLeadingBlnks(inputPairs[2]);
	 if ( ( !lastNm) || ( !frstNm)  ){ res.send('---' );   return}; // bad input
	 reslt=knownName(lastNm);
	 if ( reslt[0] == -1 ){ res.send('---' );   return};  // either not found or already some exist with first names
	 row=reslt[0];  ptr=amudot.name+row.toString();
	 requestedSeatsWorksheet[ptr].v=requestedSeatsWorksheet[ptr].v+' '+frstNm+'*';
	 xlsx.writeFile(workbook, XLSXfilename);
	 familyNames[row-startingRow]=lastNm+' '+frstNm;
	 firstName[row-startingRow]=frstNm;
	 res.send('+++');
   }
   else res.send('999' );
});
//---------------------------------------------------------------------------------   //[ulam][moed][gvarim-nashim]  assignedPerUlam  

app.get('/getCountOfSeats', function(req, res) {
	res.header("Access-Control-Allow-Origin", "*");
  
	rspns='';
	for (ulam=0; ulam<2;ulam++)      //// rashi-martef  /  gvarim-nashim
	      for (gender=0;gender<2;gender++) rspns=rspns+maxCountSeats[ulam][gender].toString()+'$';
				
	for (ulam=0; ulam<2;ulam++)             //[ulam][moed][gvarim-nashim] 
	   for (moed=0;moed<2;moed++)
		    for (gender=0;gender<2;gender++) rspns=rspns+assignedPerUlam[ulam][moed][gender].toString()+'$';	
				
	rspns=rspns.substr(0,rspns.length-1);
						
  
	 res.send(rspns);
	
	 });

//--------------------------------------------------------------------------------- maxCountSeats  assignedPerUlam  [ulam][moed][gvarim-nashim] 

// Hard initialize membersRequests.xlsx file

app.get('/shira1807', function(req, res) {
	res.header("Access-Control-Allow-Origin", "*");
	tmpfile=fs.readFileSync('EmptymembersRequests.xlsx');
	fs.writeFileSync(EmptyXLSXfilename, tmpfile);
	fs.writeFileSync(XLSXfilename, tmpfile);
	workbook = xlsx.readFile(XLSXfilename);
	requestedSeatsWorksheet = workbook.Sheets['HTMLRequests'];
	
	for(i=0;i<1500;i++){
	alreadyAssignedSeatsRosh[i]=' '; 
  alreadyAssignedSeatsKipur[i]=' ';
	 };
	 
	 initValuesOutOfHtmlRequestsXLSX_file();
	
 console.log('membersRequests file was HARD initialized');

var seatOcuupationLevel = new Array;
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
	listOfPayments='';
	for(i=firstSeatRow;i<lastSeatRow+1;i++){
	    pointerCell=amudot.name+(i).toString(); 
		 cell=requestedSeatsWorksheet[pointerCell]; 
	   if(! cell) continue;
		 listOfPayments=listOfPayments+'$'+cell.v;
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
		   if( ! paid) continue;
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

// initialize membersRequests.xlsx file

app.get('/s276662', function(req, res) {
	res.header("Access-Control-Allow-Origin", "*");
	tmpfile=fs.readFileSync(EmptyXLSXfilename);
	fs.writeFileSync(XLSXfilename, tmpfile);
	workbook = xlsx.readFile(XLSXfilename);
	requestedSeatsWorksheet = workbook.Sheets['HTMLRequests'];
	 res.setHeader('Content-Type', 'text/html'); 
	res.send(cache_get('initialize') );
	
	 
	 });

//---------------------------------------------------------------------------------	

// get request to write member's input
   var inputArray = new Array;
  app.get('/writeinfo', function(req, res) {
	res.header("Access-Control-Allow-Origin", "*");
	inputString=decodeURI(req.originalUrl);   
	
	inputString=inputString.substr(12); 
	inputPairs=inputString.split('&'); 
	namTitl=inputPairs[0].split('=')[1];
	managementRequest=false;
	sendMsgToKehilatArielSeatsGmail(namTitl,inputString);
	 if(handleInput(inputPairs)){res.send('+++'+incompatibilty+'$'+hasStillToPay.toString()+afterClosingDate)}else { res.send('---'+errorNumber);};
	 
	 });
//---------------------------------------------------------------------------------
app.get('/isThereSuchAName', function(req, res) {
  var rNmA = new Array();  
	res.header("Access-Control-Allow-Origin", "*");
	inputString=decodeURI(req.originalUrl);
	inputPairs=inputString.split('$'); 
	if (inputPairs[1] != mngmntPASSW){ res.send('999')}
	
	else { 
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
	respns=memberInfo('member',inp);
	res.send(respns);
	});
//----------------------------------------------------------------------

// get request to verify family name for mgmt

	app.get('/mngfmname', function(req, res) {
	res.header("Access-Control-Allow-Origin", "*");
	inp=decodeURI(req.originalUrl).split('?')[1];
	inpData=inp.split('-');   
	if(inpData[1] == mngmntPASSW){
	   
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


//----------------------------------------------------------
		
	
	app.get('/manage', function(req, res) {
	//res.header("Access-Control-Allow-Origin", "*");
	 res.setHeader('Content-Type', 'text/html');
            res.send(cache_get('mgmt') );
        })
				
//----------------------------------------------------------
app.get('/getFullList', function(req, res) {
	res.header("Access-Control-Allow-Origin", "*");
	 res.setHeader('Content-Type', 'text/html');
   
	 inp=decodeURI(req.originalUrl).split('?')[1];
	inpData=inp.split('$');
	 
	if(inpData[1] == mngmntPASSW){
	   
	 listOfnames='+++'+	familyNames.join('$');																														
	 res.send(listOfnames);
	}
	else {res.send('999' );
	      }
	 
	 
        })


//----------------------------------------------------------


app.get('/getlist', function(req, res) {
	res.header("Access-Control-Allow-Origin", "*");
	 res.setHeader('Content-Type', 'text/html');
   
	 inp=decodeURI(req.originalUrl).split('?')[1];
	inpData=inp.split('$');
	
	if(inpData[1] == mngmntPASSW){
	 
	 moedCode=inpData[2];
	 SidurUlam=inpData[3];
	 shlavBFilter=inpData[4];
	 notPaidFilter=inpData[5];
	 doneWithFilter=inpData[6];
	   
	 listOfnames='+++'+	filterAndSort();																														
	 res.send(listOfnames);
	}
	else res.send('999' );
	      
	 
	 
        })

//----------------------------------------------------------
 app.get('/saveStsfctionParams', function(req, res) {
	res.header("Access-Control-Allow-Origin", "*");
	 res.setHeader('Content-Type', 'text/html');
   
	 inp=decodeURI(req.originalUrl).split('?')[1];
	 inpData=inp.split('$');
	 
	if(inpData[0] == mngmntPASSW){
	   
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
		 ptr= amudot.stsfctnInFlrLastYrMen + row;
		 requestedSeatsWorksheet[ptr].v=inpData[5];   //lastYearStsfctnMen
		 ptr= amudot.stsfctnInFlrLastYrWmn + row;
		 requestedSeatsWorksheet[ptr].v=inpData[6];   //lastYearStsfctnWmn
		 ptr= amudot.lstYrSeat + row;
		 requestedSeatsWorksheet[ptr].v=inpData[7]; //lastYearSeat
		 ptr= amudot.nashimMuadaf + row;
		 requestedSeatsWorksheet[ptr].v=inpData[8]; 
		 ptr= amudot.gvarimMuadaf + row;
		 requestedSeatsWorksheet[ptr].v=inpData[9]; 
		 
		  xlsx.writeFile(workbook, XLSXfilename);																													
	   res.send('+++');
	   }
	}
	else res.send('999' );
	 
        })


				

		
//----------------------------------------------------------saveStsfctionParams


app.get('/updateAssignedSeats', function(req, res) {
	res.header("Access-Control-Allow-Origin", "*");
	 res.setHeader('Content-Type', 'text/html');
	
	 
	 inputString=decodeURI(req.originalUrl).split('?')[1];
	 inputParams=inputString.split('$'); 
	 if (inputParams[0] != mngmntPASSW){res.send('999' )}
	 else{ 
	   moed=inputParams[1]; 
	   nameToUpdate=inputParams[2];  
	   strSeatsToUpdate=inputParams[3];  
	 
	   rowNum= knownName(nameToUpdate)[0];   
		 if(rowNum != -1 ){
		 row=rowNum.toString();
		 
		 if( (moed=='rosh') || (moed=='all') ){
		 ptr=amudot.assignedSeatsRosh+row; 
		 oldAssignedSeatsRosh=requestedSeatsWorksheet[ptr].v;
		 requestedSeatsWorksheet[ptr].v=strSeatsToUpdate;  // console.log('ptr='+ptr+' requestedSeatsWorksheet[ptr].v='+requestedSeatsWorksheet[ptr].v);
		 markedSeatsLeft('rosh',row,delLeadingBlnks(oldAssignedSeatsRosh)); 
			 };
			 
		 if( (moed=='kipur') || (moed=='all') ){   
		 ptr=amudot.assignedSeatsKipur+row;
		 oldAssignedSeatsKipur=requestedSeatsWorksheet[ptr].v;
		 requestedSeatsWorksheet[ptr].v=strSeatsToUpdate;   //console.log('ptr='+ptr+' requestedSeatsWorksheet[ptr].v='+requestedSeatsWorksheet[ptr].v);
		 markedSeatsLeft('kipur',row,delLeadingBlnks(oldAssignedSeatsKipur));
		    }
		 CountAssignedPerMoed_PerUlam();
			 xlsx.writeFile(workbook, XLSXfilename);
			 res.send('+++' );
			 }
			 else res.send('---' );
			
			} 
	 
	 
	      

    })
		
//----------------------------------------------------------

function markedSeatsLeft(moed,row,oldAssignment){
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
																															break;   
																															}  // for ij
														            	requestedSeatsWorksheet[stillRequestedForMoedCol +row1].v=LeftMarkedSeatsArray.join('+');															
												      		     }  // for ii
															}  // for membr1	
															
															
	// update namesForSeat with new assigns
		for (ii=0; ii<		newlyAssignedArray.length;ii++){
		   			aSeat=Number(newlyAssignedArray[ii].split('_')[0]);
						namesForSeatParts=namesForSeat[aSeat].split('/');
            namesAssigned=namesForSeatParts[0].split('$');
			      namesAssigned[namesAssignedIdx]=requestedSeatsWorksheet[amudot.name+row].v;
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
app.get('/setdebug123', function(req, res) {
	//res.header("Access-Control-Allow-Origin", "*");
	 res.setHeader('Content-Type', 'text/html');
	 inputString=req.originalUrl.substr(13);   
	 debugprms=inputString.split('$');  
	 if(debugprms[0]=='on'){  debugIsOn=true; debugparam=debugprms[1]};
	  if(debugprms[0]=='off'){  debugIsOn=false; debugparam=''};
	
  res.send(cache_get('okmsg') );
        })	
				
			
				
//----------------------------------------------------------
	
	app.get('/initSatisfactionFile', function(req, res) {
	res.header("Access-Control-Allow-Origin", "*");
	 res.setHeader('Content-Type', 'text/html');
	
	
	
	 
	 inputString=decodeURI(req.originalUrl).split('?')[1];
	// inputParams=inputString.split('$'); 
	 if (inputString != mngmntPASSW){res.send('999' )}
	 else {
	      rtrncod=initSatisfactionFile();
				 res.send(rtrncod );
			}	  
	 })
//----------------------------------------------------------  

	
	app.get('/dnldSatisfactionFile', function(req, res) {
	res.header("Access-Control-Allow-Origin", "*");
	 res.setHeader('Content-Type', 'text/html');
	
	 inputString=decodeURI(req.originalUrl).split('-')[1];
	 if (inputString != mngmntPASSW){console.log('wrong password'); res.send('999' )}
	 else {
	       stsfctionSheets=stsfctionWB.SheetNames;
	       nm=stsfctionSheets.length;  
         nextyrSheetNum=stsfctionSheets.indexOf(crrntYrSheetName);  
         if(nextyrSheetNum== -1) stsfctionWB.SheetNames[nm]=crrntYrSheetName;
				 
				 
				 
				 stsfctionWB.Sheets[crrntYrSheetName]=stsfctionWB.Sheets['emptySheet'];
				 currentYrSheet=stsfctionWB.Sheets[crrntYrSheetName];
  				emptySheet=stsfctionWB.Sheets['emptySheet'];;
								 
				      for(i=0;i<familyNames.length  ; i++){
				          nam=familyNames[i];
				          row1=knownName(nam)[0];
									row2=row1-firstSeatRow+startingRowstsfction;
									row1=row1.toString();
									row2=row2.toString();
                  Object.keys(amudotForStsfctionDownload).forEach(function(key)  {

                            ptr1= amudotForStsfctionDownload[key][0]+row1;
									          ptr2= amudotForStsfctionDownload[key][1]+row2;
														currentYrSheet[ptr2].v=requestedSeatsWorksheet[ptr1].v;
													
													 
                         });	
									}	// for
				 		
							xlsx.writeFile(stsfctionWB, SortingDatafilename);	
							
				fileToSendName= 'seatsOrdered.xlsx'; 
				fileToRead=SortingDatafilename; 				
				res.setHeader('Content-type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
        res.setHeader('Content-Disposition', 'attachment; filename=' + fileToSendName);
	      var fileR = fs.readFileSync(fileToRead, 'binary');
        res.setHeader('Content-Length', fileR.length);
        res.write(fileR, 'binary');
        res.end();
      			
							
							
							
									
			}  // else						
								 
 
	 })
	 
	

//----------------------------------------------------------
	 
	 
function 	 initSatisfactionFile(){ 
	       errInProcess=false;
	       for(i=0;i<familyNames.length  ; i++){
				          nam=familyNames[i];
				          row1=(knownName(nam)[0]).toString();
									row2=rowNumOfStsfctionFile(nam);
									if (row2==-1){errInProcess=true; continue;};
				              row2=row2.toString();
	                     Object.keys(amudotForStsfctionUpload).forEach(function(key) {

                            ptr1= amudotForStsfctionUpload[key][0]+row1;
									          ptr2= amudotForStsfctionUpload[key][1]+row2;
														if(stsfctionWorkSheet[ptr2]){vlu=stsfctionWorkSheet[ptr2].v} else vlu=0;
														
									          requestedSeatsWorksheet[ptr1].v=vlu; 
                         });	
									// update membership status   
									ptr1=amudot.memberShipStatus+row1;
									vlu=requestedSeatsWorksheet[ptr1].v;
									if(vlu<2)requestedSeatsWorksheet[ptr1].v=vlu+1;
									
							}	// for
						
						
							
						xlsx.writeFile(workbook, XLSXfilename);		
				  retrnCode='+++'; if (	errInProcess)retrnCode='---';
					return retrnCode;
		  } 
					
													
//----------------------------------------------------------
  function rowNumOfStsfctionFile(str){
	          
	rNm=-1; 
	for (rn=0; rn<stsfctionFamilyNames.length; rn++){
	      if(stsfctionFamilyNames[rn] == str) { rNm=rn; break };
			};
			
	if(rNm!=-1)rNm=rNm+startingRowstsfction;
	return rNm;
	}; 

//----------------------------------------------------------
	 		
	//  recover data fromBackup
	
	
	
	
	app.get('/recoverBackupData2509', function(req, res) {
	//res.header("Access-Control-Allow-Origin", "*");
	 res.setHeader('Content-Type', 'text/html');
	
	 var backupWB = xlsx.readFile(BackupFilename);

	  basePayment=1200;
    paymentPerSeat=50;
		
		 
	  var backupWBSheet = backupWB.Sheets['HTMLRequests']; 
		rowNum=	firstSeatRow;
   	row=rowNum.toString();
	
		while(backupWBSheet['A'+row].v  !=  '$$$'){  console.log('row='+row);
		   Object.keys(amudot).forEach(function(key)  {

                            ptr1= amudot[key]+row;// console.log('ptr1='+ptr1);
														requestedSeatsWorksheet[ptr1].v = backupWBSheet[ptr1].v ;
						      });	

					// go to next name
			  		rowNum++;
						row=rowNum.toString();  console.log('row='+row);
			}			

					// all done -	write file 
					xlsx.writeFile(workbook, XLSXfilename); 
					
					initValuesOutOfHtmlRequestsXLSX_file();
					  
            res.send(cache_get('initialize') );
        })					




//----------------------------------------------------------
	 		
	//   set input data to last year
	
	
	
	
	app.get('/set2015data2509', function(req, res) {
	//res.header("Access-Control-Allow-Origin", "*");
	 res.setHeader('Content-Type', 'text/html');
	
	 var SeatPrio2015WB = xlsx.readFile('SeatPrio2015a.xlsx');

	  basePayment=1200;
    paymentPerSeat=50;
		
		 
	 var SeatPrio2015WBSheet = SeatPrio2015WB.Sheets['requestedSeats1']; 
	  namePtr='F7';
		ptrToRequestCol=namePtr.substr(0,namePtr.length-1); 
		name2015=SeatPrio2015WBSheet[namePtr].v; 
		while(name2015 != '$$$'){
		   rowNum=knownName(name2015)[0];
			 if (rowNum != -1){
			     seatStr='';
					 wmn=0;
					 men=0;
		
		       for (i=9;i<624;i++){
					  ptrToRequest=ptrToRequestCol+(i).toString();
						if (SeatPrio2015WBSheet[ptrToRequest].v != 1)continue;
						seatStr=seatStr+'+'+SeatPrio2015WBSheet['D'+(i).toString()].v+'_1';
						if(SeatPrio2015WBSheet['B'+(i).toString()].v == 1){wmn++;}else men++;
	             } // end loop on i
						if(seatStr){  // at least one seat selected	
						  //write info 
						    requestedSeatsWorksheet[amudot.markedSeats+rowNum.toString()].v=seatStr.substr(1);		
								requestedSeatsWorksheet[amudot.notAssignedMarkedSeatsRosh+rowNum.toString()].v=seatStr.substr(1);
						    requestedSeatsWorksheet[amudot.notAssignedMarkedSeatsKipur+rowNum.toString()].v=seatStr.substr(1);
								
								requestedSeatsWorksheet[amudot.assignedSeatsKipur+rowNum.toString()].v='';
								requestedSeatsWorksheet[amudot.assignedSeatsRosh+rowNum.toString()].v='';
								requestedSeatsWorksheet[amudot.numberOfAssignedSeatsRoshMen+rowNum.toString()].v='0';
								requestedSeatsWorksheet[amudot.numberOfAssignedSeatsRoshWomen+rowNum.toString()].v='0';
								requestedSeatsWorksheet[amudot.numberOfAssignedSeatsKipurMen+rowNum.toString()].v='0';
								requestedSeatsWorksheet[amudot.numberOfAssignedSeatsKipurWomen+rowNum.toString()].v='0';
								ptrN=amudot.numberMarkedMen+rowNum.toString();  requestedSeatsWorksheet[ptrN].v=men.toString();
								ptrN=amudot.numberMarkedWomen+rowNum.toString();  requestedSeatsWorksheet[ptrN].v=wmn.toString();
								requestedSeatsWorksheet[amudot.NumberOfNotAssignedMarkedSeatsMen+rowNum.toString()].v=men.toString();
								requestedSeatsWorksheet[amudot.NumberOfNotAssignedMarkedSeatsWomen+rowNum.toString()].v=wmn.toString();
								
								ptrZ=ptrToRequestCol+'4';
								ptrN=amudot.menRosh+rowNum.toString();  requestedSeatsWorksheet[ptrN].v=SeatPrio2015WBSheet[ptrZ].v;
								seatsForRoshHashana=Number(requestedSeatsWorksheet[ptrN].v);
								ptrZ=ptrToRequestCol+'3';
								ptrN=amudot.womenRosh+rowNum.toString();  requestedSeatsWorksheet[ptrN].v=SeatPrio2015WBSheet[ptrZ].v;
								seatsForRoshHashana=seatsForRoshHashana+Number(requestedSeatsWorksheet[ptrN].v);
								ptrZ=ptrToRequestCol+'5';
								ptrN=amudot.womenKipur+rowNum.toString();  requestedSeatsWorksheet[ptrN].v=SeatPrio2015WBSheet[ptrZ].v;
								seatsForYomKipur=Number(requestedSeatsWorksheet[ptrN].v);
								ptrZ=ptrToRequestCol+'6';
								ptrN=amudot.menKipur+rowNum.toString();  requestedSeatsWorksheet[ptrN].v=SeatPrio2015WBSheet[ptrZ].v;
								seatsForYomKipur=seatsForYomKipur+Number(requestedSeatsWorksheet[ptrN].v);
								
								ptrN=amudot.preferedMinyanM+rowNum.toString();  requestedSeatsWorksheet[ptrN].v='1';
								ptrN=amudot.preferedMinyanW+rowNum.toString();  requestedSeatsWorksheet[ptrN].v='1';
								
							  numberOfSeatsForPayment=Math.max(seatsForRoshHashana,	seatsForYomKipur);  
		            requiredPayment= basePayment+		numberOfSeatsForPayment * paymentPerSeat; 
								ptrN=amudot.tashlum+rowNum.toString();  requestedSeatsWorksheet[ptrN].v=requiredPayment;
								
								ptrN=amudot.assignedSeatsRosh+rowNum.toString();  requestedSeatsWorksheet[ptrN].v=' ';
								ptrN=amudot.assignedSeatsKipur+rowNum.toString();  requestedSeatsWorksheet[ptrN].v=' ';
								
								ptrN=amudot.tashlumPaid+rowNum.toString();  requestedSeatsWorksheet[ptrN].v=' ';

								 
								
								
							} // info for this seat was written
					} // end if rownum != -1
					// go to next name
			  		 
						namePtr=	SeatPrio2015WBSheet[ptrToRequestCol+'2'].v;  
						ptrToRequestCol=namePtr.substr(0,namePtr.length-1);  
						
						name2015=SeatPrio2015WBSheet[namePtr].v;  
						} //  go to next name 
					// all done -	write file 
					xlsx.writeFile(workbook, XLSXfilename);   
            res.send(cache_get('initialize') );
        })													
//----------------------------------------------------------
		
	
	app.get('/gizbar', function(req, res) {
	//res.header("Access-Control-Allow-Origin", "*");
	 res.setHeader('Content-Type', 'text/html');
            res.send(cache_get('gizbar') );
        })						
//-----------------------------------------------------------
app.get('/prtMartef', function(req, res) {
	//res.header("Access-Control-Allow-Origin", "*");
	 res.setHeader('Content-Type', 'text/html');
	 inputString=decodeURI(req.originalUrl);  
	 if(inputString.split('&')[1] == mngmntPASSW){  res.send(cache_get('prtMartef') )   }
	  else res.send(cache_get('errPassw') );
        })			
				
//--------------------------------------------------------------
app.get('/prtRashi', function(req, res) {
	//res.header("Access-Control-Allow-Origin", "*");
	 res.setHeader('Content-Type', 'text/html');
	 inputString=decodeURI(req.originalUrl);  
	 if(inputString.split('&')[1] == mngmntPASSW){  res.send(cache_get('prtRashi') )   }
	  else res.send(cache_get('errPassw') );
            
        })			
				
//-------------------------------------------------------

app.get('/prtNashim', function(req, res) {
	//res.header("Access-Control-Allow-Origin", "*");
	 res.setHeader('Content-Type', 'text/html');
	 inputString=decodeURI(req.originalUrl);  
	 if(inputString.split('&')[1] == mngmntPASSW){  res.send(cache_get('prtNashim') )   }
	  else res.send(cache_get('errPassw') );
	
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

// ck if registration is closed


app.get('/isRegistrationClosed', function(req, res) {
	res.header("Access-Control-Allow-Origin", "*");
	 res.setHeader('Content-Type', 'text/html');
	 isWomanString=isWoman.join('+');
	 rspns='---$ $'+isWomanString;
	 ptr=amudot.registrationClosedDateNTime+'2';  
	 tmp=delLeadingBlnks(requestedSeatsWorksheet[ptr].v); 
	 if(tmp)rspns='+++$'+tmp+'$'+isWomanString;
	 res.send(rspns );
        })	

//------------------------------------------------------------------------------					
				

app.get('/ckpswGIZBAR', function(req, res) {
	res.header("Access-Control-Allow-Origin", "*");
	 res.setHeader('Content-Type', 'text/html');
	 
	 inputString=req.originalUrl.substr(14);
	 if(inputString==gizbarPASSW){rspns='+++';} else rspns='---';
            res.send(rspns );
        })	

//------------------------------------------------------------------------------					
				
	app.get('/', function(req, res) {
	//res.header("Access-Control-Allow-Origin", "*");
	 res.setHeader('Content-Type', 'text/html');
            res.send(cache_get('index.html') );
        })			
			
				start();
				
				
				
				
	