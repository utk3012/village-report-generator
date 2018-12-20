var XLSX = require('xlsx');
var Jimp = require("jimp");
var fs = require('fs');
var createHtml = require('create-html');

var workbook = XLSX.readFile('responses3.xlsx', {cellDates: true});
var sheet_name_list = workbook.SheetNames;
var xlData = XLSX.utils.sheet_to_json(workbook.Sheets[sheet_name_list[0]]);
var columnNames = XLSX.utils.sheet_to_json(workbook.Sheets[sheet_name_list[0]], {header: 1})[0];

if (!fs.existsSync('./output-htmls')){
    fs.mkdirSync('./output-htmls');
}
if (!fs.existsSync('./output-images')){
    fs.mkdirSync('./output-images');
}

function GetFormattedDate(ddd) {
    dd = new Date(ddd);
    var month = dd.getMonth() + 1;
    var day = dd.getDate();
    var year1 = dd.getFullYear();
    if (month < 10) {
      month = '0' + month;
    }
    if (day < 10) {
      day = '0' + day;
    }
  return day + "-" + month + "-" + year1;
}

async function loadImage(cap1, cap2, cap3, cap4, cap5, cap6, cap7, cap8, cap9, cap10, cap11, cap12, cap13, cap14, cap15, cap16, cap17, cap18, cap19, cap20, availability, fileName) {
  try {
    const image = await Jimp.read('water.png');
    const font = await Jimp.loadFont(Jimp.FONT_SANS_32_BLACK);
    const font1 = await Jimp.loadFont('fonts/greyfont.fnt');
    const font2 = await Jimp.loadFont('fonts/redfont.fnt');
    const font3 = await Jimp.loadFont(Jimp.FONT_SANS_16_BLACK);
          image
          .print(font, 549, 50, cap1)
          .print(font, 743, 130, cap2)
          .print(font, 920, 135, cap3)
          .print(font, 885, 357, cap4)
          .print(font, 650, 362, cap5)
          .print(font, 359, 315, cap6)
          .print(font, 417, 420, cap7)
          .print(availability[0] ? font1 : font2, 60, 390, cap8)
          .print(availability[1] ? font1 : font2, 100, 390, cap9)
          .print(availability[2] ? font1 : font2, 140, 390, cap10)
          .print(availability[3] ? font1 : font2, 182, 390, cap11)
          .print(availability[4] ? font1 : font2, 60, 413, cap12)
          .print(availability[5] ? font1 : font2, 100, 413, cap13)
          .print(availability[6] ? font1 : font2, 140, 413, cap14)
          .print(availability[7] ? font1 : font2, 182, 413, cap15)
          .print(availability[8] ? font1 : font2, 60, 436, cap16)
          .print(availability[9] ? font1 : font2, 100, 436, cap17)
          .print(availability[10] ? font1 : font2, 140, 436, cap18)
          .print(availability[11] ? font1 : font2, 182, 436, cap19)
          .print(font3, 60, 465, cap20)
            .write(fileName);
  }
  catch (err) {
    console.error(err);
  }
}

for (var i=0; i<183; i++) {
	var year = xlData[i][columnNames[19]];
	var commitee = String(xlData[i][columnNames[41]]);
  
  if (commitee === 'Other (please specify)') {
    commitee = xlData[i][columnNames[42]];
  }
  year = String(year).substring(11,15);

  var femaleAtt = 'N/A';
  if (String(xlData[i][columnNames[95]]) !== '' && String(xlData[i][columnNames[97]]) !== '') {
    var femaleAtt1 = Number(xlData[i][columnNames[95]]);
	  var femaleAtt2 = Number(xlData[i][columnNames[97]]);
    if (femaleAtt1 + femaleAtt2 === 0) {
      femaleAtt = 0;
    }
    else {  
	   femaleAtt = ((femaleAtt1 / (femaleAtt1 + femaleAtt2))*100).toFixed(2);
    }
  }

	var registers = xlData[i][columnNames[75]];

	var netIncome = 'N/A';
	if (String(xlData[i][columnNames[126]]) !== '' && String(xlData[i][columnNames[112]]) !== '')
		var netIncome = Number(xlData[i][columnNames[126]]) - Number(xlData[i][columnNames[112]]);
	
	var regist = 0;
	if (String(xlData[i][columnNames[75]]) !== '')
		regist = String(xlData[i][columnNames[75]]).split(',').length;

  var wssSetup = xlData[i][columnNames[281]];
  wssSetup = 2018 - Number(wssSetup.toISOString().split('-')[0]);
  wssSetup = wssSetup + (wssSetup > 1 ? ' years' : 'year');

  var cfund1 = xlData[i][columnNames[243]];
  var cfund2 = xlData[i][columnNames[245]];
  var cfund3 = xlData[i][columnNames[247]];
  var cfund4 = xlData[i][columnNames[249]];
  var cfund = '';
  if (String(cfund1) !== '' && Number(cfund1) === 1) {
    cfund += 'building new toilet; '
  }
  if (String(cfund2) !== '' && Number(cfund1) === 1) {
    cfund += 'building new bathroom; '
  }
  if (String(cfund3) !== '' && Number(cfund1) === 1) {
    cfund += 'building new TBR; '
  }
  if (String(cfund4) !== '' && Number(cfund1) === 1) {
    cfund += 'WSS Maintenance; '
  }
  if(cfund === '') {
    cfund = 'None';
  }

  var MaturityDate = 'N/A';
  if (String(xlData[i][columnNames[267]]) !== '')
    MaturityDate = GetFormattedDate(xlData[i][columnNames[267]].toISOString());

	var tim1 = xlData[i][columnNames[17]];
	var tim2 = xlData[i][columnNames[18]];
	var tim3 = xlData[i][columnNames[19]];
	var tim4 = xlData[i][columnNames[44]];
	var tim5 = xlData[i][columnNames[280]];
	var tim6 = xlData[i][columnNames[281]];

	var subsImg1 = '';
	var waterSourceImg = '';
	var ia = 0;
	var ig = 0;

  if (String(xlData[i][columnNames[234]]) !== '')
    ia = Number(xlData[i][columnNames[234]]);
  if (String(xlData[i][columnNames[240]]) !== '')
    ig = Number(xlData[i][columnNames[240]]);

  if (String(xlData[i][columnNames[304]]) !== '')
    subsImg1 = `<img src="${xlData[i][columnNames[304]]}">`;

  var sourceCode = 0;
  /*
  Gravity Spring 1: 1
  Borewell 1: 2
  Borewell 2: 3
  Dugwell 1: 4
  Dugwell 2: 5
  Could not define: 6
  */

	if (String(xlData[i][columnNames[493]]) !== '') {
		waterSourceImg = `<img src="${xlData[i][columnNames[493]]}">`;
    sourceCode = 1;
  }
  else if (String(xlData[i][columnNames[368]]) !== '') {
    waterSourceImg = `<img src="${xlData[i][columnNames[368]]}">`; 
    sourceCode = 2;
  }
  else if (String(xlData[i][columnNames[402]]) !== '') {
    waterSourceImg = `<img src="${xlData[i][columnNames[402]]}">`; 
    sourceCode = 3;
  }
  else if (String(xlData[i][columnNames[433]]) !== '') {
    waterSourceImg = `<img src="${xlData[i][columnNames[433]]}">`; 
    sourceCode = 4;
  }
  else if (String(xlData[i][columnNames[464]]) !== '') {
    waterSourceImg = `<img src="${xlData[i][columnNames[464]]}">`;
    sourceCode = 5;
  }
  else {
    waterSourceImg = '';
    sourceCode = 6;
  }

  /* ============================================= Water Source functional logic ================================================== */
  var waterStatus = '';
  var outputStatement = '';
  var waterStatusAcToCommitee = String(xlData[i][columnNames[296]]); // Condition 1
  var houseHoldsNotConnected = String(xlData[i][columnNames[334]]);  // Condition 3
  var householdsNotReciveing = String(xlData[i][columnNames[338]]);  // Condition 3

  var pump = '';
  var pumpStatus = '';

  if (sourceCode === 1) { //means gravity spring 1
    waterStatus = String(xlData[i][columnNames[500]]);
    var noOfMonths = waterStatus.split(',').length;
    var householdsExcluded = Number(houseHoldsNotConnected) + Number(householdsNotReciveing);
    if (householdsExcluded === 0)
      householdsExcluded = 'no';
    if (waterStatusAcToCommitee === 'Yes' && houseHoldsNotConnected === '0' && householdsNotReciveing === '0' &&
      waterStatus === 'Water available throughout year') {
        outputStatement = ' is <u>Functional</u>, water is supplied to all households throughout the year.';
    }
    else if (waterStatusAcToCommitee === 'No' && waterStatus !== 'Water available throughout year' && waterStatus !== '' && noOfMonths >= 6) {
      outputStatement = ` is <u>Non-functional</u>, water not supplied during ${noOfMonths} months and ${householdsExcluded} households excluded.`;
    }
    else {
      outputStatement = ` is <u>Partly functional</u>, water not supplied during ${noOfMonths} ` + (noOfMonths === 1 ? 'month' : 'months') + ` and ${householdsExcluded} households excluded.`; 
    }
  }
  else if (sourceCode === 2) { //means borewell 1
    waterStatus = String(xlData[i][columnNames[378]]);
    pump = String(xlData[i][columnNames[373]]);
    pumpStatus = String(xlData[i][columnNames[394]]);
    var noOfMonths = waterStatus.split(',').length;
    var householdsExcluded = Number(houseHoldsNotConnected) + Number(householdsNotReciveing);
    if (householdsExcluded === 0)
      householdsExcluded = 'no';
    if (waterStatusAcToCommitee === 'Yes' && houseHoldsNotConnected === '0' && householdsNotReciveing === '0' &&
      waterStatus === 'Water available throughout year' && pump === 'Yes' && pumpStatus === 'Running') {
        outputStatement = ' is <u>Functional</u>, water is supplied to all households throughout the year.';
    }
    else if (waterStatusAcToCommitee === 'No' && waterStatus !== 'Water available throughout year' && waterStatus !== '' && noOfMonths >= 6 && pump === 'Yes') {
      outputStatement = ` is <u>Non-functional</u>, water not supplied during ${noOfMonths} months and ${householdsExcluded} households excluded.`;
    }
    else {
      outputStatement = ` is <u>Partly functional</u>, water not supplied during ${noOfMonths} ` + (noOfMonths === 1 ? 'month' : 'months') + ` and ${householdsExcluded} households excluded.`; 
    }
  }
  else if (sourceCode === 3) { //means borewell 2
    waterStatus = String(xlData[i][columnNames[412]]);
    pump = String(xlData[i][columnNames[407]]);
    pumpStatus = String(xlData[i][columnNames[426]]);
    var noOfMonths = waterStatus.split(',').length;
    var householdsExcluded = Number(houseHoldsNotConnected) + Number(householdsNotReciveing);
    if (householdsExcluded === 0)
      householdsExcluded = 'no';
    if (waterStatusAcToCommitee === 'Yes' && houseHoldsNotConnected === '0' && householdsNotReciveing === '0' &&
      waterStatus === 'Water available throughout year' && pump === 'Yes' && pumpStatus === 'Running') {
        outputStatement = ' is <u>Functional</u>, water is supplied to all households throughout the year.';
    }
    else if (waterStatusAcToCommitee === 'No' && waterStatus !== 'Water available throughout year' && waterStatus !== '' && noOfMonths >= 6 && pump === 'Yes') {
      outputStatement = ` is <u>Non-functional</u>, water not supplied during ${noOfMonths} months and ${householdsExcluded} households excluded.`;
    }
    else {
      outputStatement = ` is <u>Partly functional</u>, water not supplied during ${noOfMonths} ` + (noOfMonths === 1 ? 'month' : 'months') + ` and ${householdsExcluded} households excluded.`; 
    }
  }
  else if (sourceCode === 4) { //means dugwell 1
    waterStatus = String(xlData[i][columnNames[441]]);
    pump = String(xlData[i][columnNames[434]]);
    pumpStatus = String(xlData[i][columnNames[457]]);
    var noOfMonths = waterStatus.split(',').length;
    var householdsExcluded = Number(houseHoldsNotConnected) + Number(householdsNotReciveing);
    if (householdsExcluded === 0)
      householdsExcluded = 'no';
    if (waterStatusAcToCommitee === 'Yes' && houseHoldsNotConnected === '0' && householdsNotReciveing === '0' &&
      waterStatus === 'Water available throughout year' && pump === 'Yes' && pumpStatus === 'Running') {
        outputStatement = ' is <u>Functional</u>, water is supplied to all households throughout the year.';
    }
    else if (waterStatusAcToCommitee === 'No' && waterStatus !== 'Water available throughout year' && waterStatus !== '' && noOfMonths >= 6 && pump === 'Yes') {
      outputStatement = ` is <u>Non-functional</u>, water not supplied during ${noOfMonths} months and ${householdsExcluded} households excluded.`;
    }
    else {
      outputStatement = ` is <u>Partly functional</u>, water not supplied during ${noOfMonths} ` + (noOfMonths === 1 ? 'month' : 'months') + ` and ${householdsExcluded} households excluded.`; 
    }
  }
  else if (sourceCode === 5) { //means dugwell 2
    waterStatus = String(xlData[i][columnNames[472]]);
    pump = String(xlData[i][columnNames[465]]);
    pumpStatus = String(xlData[i][columnNames[486]]);
    var noOfMonths = waterStatus.split(',').length;
    var householdsExcluded = Number(houseHoldsNotConnected) + Number(householdsNotReciveing);
    if (householdsExcluded === 0)
      householdsExcluded = 'no';
    if (waterStatusAcToCommitee === 'Yes' && houseHoldsNotConnected === '0' && householdsNotReciveing === '0' &&
      waterStatus === 'Water available throughout year' && pump === 'Yes' && pumpStatus === 'Running') {
        outputStatement = ' is <u>Functional</u>, water is supplied to all households throughout the year.';
    }
    else if (waterStatusAcToCommitee === 'No' && waterStatus !== 'Water available throughout year' && waterStatus !== '' && noOfMonths >= 6 && pump === 'Yes') {
      outputStatement = ` is <u>Non-functional</u>, water not supplied during ${noOfMonths} months and ${householdsExcluded} households excluded.`;
    }
    else {
      outputStatement = ` is <u>Partly functional</u>, water not supplied during ${noOfMonths} ` + (noOfMonths === 1 ? 'month' : 'months') + ` and ${householdsExcluded} households excluded.`; 
    }
  }
  else {
    waterStatus = '';
    outputStatement = 'water source could be defined';
  }

  // console.log(outputStatement);

  /* =============================================End of Water Source functional logic ================================================== */

  /* ============================================= Water Source count statement calculation ================================================== */

  var borewellCount = String(xlData[i][columnNames[361]]);
  var dugwellCount = String(xlData[i][columnNames[362]]);
  var gravitySpringCount = String(xlData[i][columnNames[363]]);
  var borewellStatement = '';
  var dugwellStatement = '';
  var gravitySpringStatement = '';
  var outputCountStatement = '';

  if (borewellCount !== '' && Number(borewellCount) !== 0) {
    borewellStatement += borewellCount + (Number(borewellCount) > 1 ? ' Borewells' : ' Borewell');
  }
  if (dugwellCount !== '' && Number(dugwellCount) !== 0) {
    dugwellStatement += dugwellCount + (Number(dugwellCount) > 1 ? ' Dugwells' : ' Dugwell');
  }
  if (gravitySpringCount !== '' && Number(gravitySpringCount) !== 0) {
    gravitySpringStatement += gravitySpringCount + (Number(gravitySpringCount) > 1 ? ' Gravity Springs' : ' Gravity Spring');
  }
  if (gravitySpringStatement !== '' && borewellStatement !== '' && dugwellStatement !== '') {
    outputCountStatement = borewellStatement + ', ' + dugwellStatement + ' and ' + gravitySpringStatement + ' are used for Water supply system.';
  }
  if (gravitySpringStatement !== '' && borewellStatement !== '' && dugwellStatement === '') {
    outputCountStatement = borewellStatement + ' and ' + gravitySpringStatement + ' are used for Water supply system.';
  }
  if (gravitySpringStatement !== '' && borewellStatement === '' && dugwellStatement !== '') {
    outputCountStatement = dugwellStatement + ' and ' + gravitySpringStatement + ' are used for Water supply system.';
  }
  if (gravitySpringStatement === '' && borewellStatement !== '' && dugwellStatement !== '') {
    outputCountStatement = borewellStatement + ' and ' + dugwellStatement + 'are used for Water supply system.';
  }
  if (gravitySpringStatement !== '' && borewellStatement === '' && dugwellStatement === '') {
    outputCountStatement = gravitySpringStatement + (Number(gravitySpringCount) > 1 ? ' are' : ' is') + ' used for Water supply system.';
  }
  if (gravitySpringStatement === '' && borewellStatement !== '' && dugwellStatement === '') {
    outputCountStatement = borewellStatement + (Number(borewellCount) > 1 ? ' are' : ' is') + ' used for Water supply system.';
  }
  if (gravitySpringStatement === '' && borewellStatement === '' && dugwellStatement !== '') {
    outputCountStatement = dugwellStatement + (Number(dugwellCount) > 1 ? ' are' : ' is') + ' used for Water supply system.';
  }

  // console.log(outputCountStatement);

  /* ============================================= End of Water Source count statement calculation =========================================*/

  /* ============================================= Image value retrival ==============================================*/

  var kv = 'NA';
  var ll = 'NA';
  var an = 'NA';
  var of = String(xlData[i][columnNames[395]]);
  var qo = String(xlData[i][columnNames[456]]);
  var lw = xlData[i][columnNames[334]];
  var ma = xlData[i][columnNames[338]];
  var lwma = Number(lw) + Number(ma);
  
  var mb = xlData[i][columnNames[339]];
  var mc = xlData[i][columnNames[340]];
  var mbmc = '';
  if (String(mb) === 'Other (please specify)') {
    mbmc = String(mc);
  }
  else if(String(mb) === '' && String(mc) === '') {
    mbmc = '';
  }
  else {
    mbmc = String(mb);
  }

  var mi = String(xlData[i][columnNames[346]]);
  var mk = '';
  if (mi === 'Yes') {
    mk = 'None';
    if (String(xlData[i][columnNames[348]]) !== '') {
      mk = String(xlData[i][columnNames[348]]);
    }
    if (String(xlData[i][columnNames[348]]) !== 'Other (please specify)') {
      mk = String(xlData[i][columnNames[349]]);
    }
    mi = mi + ', ' + mk;
  }
  
  if (String(xlData[i][columnNames[307]]) !== '')
    kv = Number(xlData[i][columnNames[307]]);
  if (String(xlData[i][columnNames[323]]) !== '')
    ll = Number(xlData[i][columnNames[323]]);
  if (String(xlData[i][columnNames[39]]) !== '')
    an = Number(xlData[i][columnNames[39]]);
  if (of !== '')
    of = Number(xlData[i][columnNames[395]]);
  else
    of = 0;
  if (qo !== '')
    qo = Number(xlData[i][columnNames[456]]);
  else
    qo = 0;

  var monthsName = ['January', 'February', 'March', 'April', 'May', 'June', 'July', 'August', 'September', 'October', 'November', 'December'];
  var availability = [true, true, true, true, true, true, true, true, true, true, true, true];
  var availableThroughout = '';
  if (waterStatus === 'Water available throughout year') {
    availableThroughout = waterStatus;
  }
  else {
    for (var j = 0; j < 12; j++) {
      if (waterStatus.indexOf(monthsName[j]) !== -1) {
        availability[j] = false;
      }
    }
  }

  loadImage(kv, ll, an, lwma, lw, String(of+qo), mi, 'JAN', 'FEB', 'MAR', 'APR', 'MAY', 'JUN', 'JUL', 'AUG', 'SEP', 'OCT', 'NOV', 'DEC', availableThroughout, availability, `output-images/water${i+1}.png`);

  /* ============================================= End of Image value retrival ===========================================*/
	

	var html = createHtml({
	  title: `${xlData[i][columnNames[16]]}`,
	  lang: 'en',
	  css: '../style.css',
	  head: `<meta name="description" content="report">
    <link rel="stylesheet" type="text/css" href="https://cdnjs.cloudflare.com/ajax/libs/semantic-ui/2.4.1/components/table.min.css">
    <script src="https://cdnjs.cloudflare.com/ajax/libs/vis/4.21.0/vis.min.js"></script>
    <link href="https://cdnjs.cloudflare.com/ajax/libs/vis/4.21.0/vis.min.css" rel="stylesheet" type="text/css" />`,
    body: `
   <div id="container">
      <div class="center">
        <span id="header1">
            <strong>Village Water and Sanitation Status Assessment Report</strong> : ${xlData[i][columnNames[16]]}
        </span>
        <br>
        <span id="header2">${xlData[i][columnNames[15]]} Gram Panchayat, ${xlData[i][columnNames[14]]} Block, ${xlData[i][columnNames[13]]} District</span>
        <br>
        <span style="font-size: 17px;">Completed in ${year} &emsp; | &emsp; ${xlData[i][columnNames[20]]} Households in ${year} &emsp; | &emsp; ${xlData[i][columnNames[21]]} Households in 2018</span>
      </div>
    <hr>
    <span class="header4"><strong>Water supply is ${wssSetup} old; ${outputStatement}</strong></span>
    <br>
    <span class="header3">${outputCountStatement}</span>
    <br>
    <br>
    <div id="imgDiv">
      <img src="../output-images/water${i+1}.png" height="495px" width="1030px" style="border: 1px solid black">
    </div>
    <div style="clear: both;"></div>
    <h4 class="header4" style="margin-bottom: 5px; margin-top: 14px;">Timeline</h4>
    <div id="visualization" style="font-size: 12px"></div>
    <br>
    <span class="header4"><strong>Village Commitee - </strong>${commitee} <strong>|</strong>  ${xlData[i][columnNames[57]]}</span>
    <br><br>
    <div id="tableData">
    <table class="ui celled selectable small compact table" style="font-size: 13px">
  <tbody>
    <tr class="center aligned">
      <td><strong>${xlData[i][columnNames[40]] || 'N/A'} </strong><br>Commitee exists</td>
      <td><strong>${xlData[i][columnNames[54]] || '0'} </strong><br>VWSC members</td>
      <td><strong>${xlData[i][columnNames[56]] || '0'} </strong><br>Women VWSC members</td>
      <td><strong>${xlData[i][columnNames[68]] || '0'} </strong><br>VWSC leadership change</td>
      <td><strong>${xlData[i][columnNames[52]] || '0'} </strong><br>Meetings in 12 months</td>
      <td><strong>${xlData[i][columnNames[51]] || '0'} </strong><br>Meetings in register</td>
    </tr>
    <tr class="center aligned">
      <td><strong>${xlData[i][columnNames[43]] || 'N/A'} </strong><br>Committee registration</td>
      <td><strong>${xlData[i][columnNames[161]] || '0'} </strong><br>Maintenance fund</td>
      <td><strong>${xlData[i][columnNames[233]] || 'N/A'} </strong><br>Corpus fund</td>
      <td><strong>${xlData[i][columnNames[143]] || 'N/A'} </strong><br>PAN</td>
      <td><strong>${regist} </strong><br>Registers</td>
      <td><strong>${femaleAtt} %  </strong><br>Female Attendance</td>
    </tr>
  </tbody>
</table>
<div id="maint" style="margin-top: -20px">
<h4>Maintenance Finances</h4>
    <div id="tableData">
    <table class="ui celled selectable small compact table" style="font-size: 13px">
  <tbody>
    <tr>
      <td colspan="2" style="text-transform: uppercase;"><strong>Monthly Revenue</strong></td>
    </tr>
    <tr>
      <td>User Fee per household</td>
      <td><strong>${xlData[i][columnNames[208]] || '0'}</strong></td>
    </tr>
    <tr>
      <td>Total Revenues</td>
      <td><strong>${xlData[i][columnNames[126]] || '0'}</strong></td>
    </tr>
    <tr>
      <td colspan="2" style="text-transform: uppercase;"><strong>Monthly Expenses</strong></td>
    </tr>
    <tr>
      <td>Electricity charges</td>
      <td><strong>${xlData[i][columnNames[124]] || '0'}</strong></td>
    </tr>
    <tr>
      <td>Operator salary in cash</td>
      <td><strong>${xlData[i][columnNames[110]] || '0'}</strong></td>
    </tr>
    <tr>
      <td>Total Expenses</td>
      <td><strong>${xlData[i][columnNames[112]] || '0'}</strong></td>
    </tr>
    <tr>
      <td style="color: black;">Net Monthly Income</td>
      <td><strong>${netIncome}</strong></td>
    </tr>
    </tbody>
</table>
</div>
</div>
<div id="corpus" style="margin-top: -20px">
<h4>Corpus fund</h4>
    <div id="tableData">
    <table class="ui celled selectable small compact table" style="font-size: 13px">
  <tbody>
    <tr>
      <td>Village has corpus fund?</td>
      <td><strong>${xlData[i][columnNames[233]] || 'N/A'}</strong></td>
    </tr>
    <tr>
      <td>Corpus fund amount</td>
      <td><strong>${xlData[i][columnNames[235]] || 'N/A'}</strong></td>
    </tr>
    <tr>
      <td>Households contributed</td>
      <td><strong>${ia+ig}</strong></td>
    </tr>
    <tr>
      <td>No of times corpus fund was used</td>
      <td><strong>${xlData[i][columnNames[242]] || 'N/A'}</strong></td>
    </tr>
    <tr>
      <td>Corpus funds used for</td>
      <td><strong>${cfund}</strong></td>
    </tr>
    <tr>
      <td>Maturity Date</td>
      <td><strong>${MaturityDate}</strong></td>
    </tr>
    <tr>
      <td>Maturity amount</td>
      <td><strong>${xlData[i][columnNames[263]] || 'N/A'}</strong></td>
    </tr>
    <tr>
      <td>Interest Rate</td>
      <td><strong>${xlData[i][columnNames[265]] || 'N/A'}</strong></td>
    </tr>
    </tbody>
</table>
</div>
</div>
<br>
<div class="subsystem">
  <div class="subsystemTable" style="margin-top: 15px;">
  <table class="ui celled table compact center aligned">
    <tbody>
      <tr>
        <td>
        <div>
    <strong>Village location</strong> <br>
      <img id="map" src="https://maps.googleapis.com/maps/api/staticmap?zoom=9&size=300x200&maptype=roadmap&markers=color:red%7C${xlData[i][columnNames[8]]},${xlData[i][columnNames[9]]}&key=AIzaSyBHJrfsTYyp2YsABs2YPmzYTpfI6SoPqv4">
  </div>
  </td>
  <td>
  <div>
    <strong>Water tank </strong><br>
    ${subsImg1}
  </div>
  </td>
  <td>
  <div>
    <strong>Water source </strong><br>
    ${waterSourceImg}
  </div>
  </td></tr>
</tbody>
</table>
</div>
</div>

<script type="text/javascript">
  // DOM element where the Timeline will be attached
  var container = document.getElementById('visualization');

  // Create a DataSet (allows two way data-binding)
  var items = new vis.DataSet([
    {id: 1, content: 'GV first visited', start: '${tim1}'},
    {id: 2, content: 'TBR work began', start: '${tim2}'},
    {id: 3, content: 'TBR 100% complete', start: '${tim3}'},
    {id: 4, content: 'Commitee formed', start: '${tim4}'},
    {id: 5, content: 'WSS work began', start: '${tim5}'},
    {id: 6, content: 'WSS work complete', start: '${tim6}'}
  ]);

  // Configuration for the Timeline
  var options = {height: '160px'};

  // Create a Timeline
  var timeline = new vis.Timeline(container, items, options);
	</script>
</div>
</div>`
	});


	fs.writeFile(`output-htmls/index${i+1}.html`, html, function (err) {
	  if (err) console.log(err);
	  console.log(`Report generated`);
	});

}