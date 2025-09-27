// Google Apps Script (Code.gs)

var SPREADSHEET_ID = '1cWB2BNHvKXm9jIjTOFxpDREMSvHj5JCZ6r-PI4a6GeU';
var SHEET_NAME = 'Data';

/**
 * Serves the dashboard + forms HTML
 */
function doGet(e) {
  return HtmlService.createTemplateFromFile('Index')
    .evaluate()
    .setTitle('Hospital Day Reports Dashboard')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

/**
 * Allows including HTML partials
 */
function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

/**
 * Handles OPD form submission
 * @param {{date:string, numberOfOPDVisit:number, numberOfNebulization:number, numberOfInjections:number, numberOfDressing:number, numberOfMinorSurgeries:number, numberOfOPDUmandawa:number, numberOfBloodDrawing:number, numberOfCatheterization:number}} formData
 */

function submitOPD(formData) {
  var dateString = formData.date;
  var row = getOrCreateRow(dateString);
  var sheet = getSheet();
  var headerMap = getHeaderMap(sheet);

  var timestamp = new Date();
  sheet.getRange(row, headerMap['OPD Timestamp - OPD']).setValue(timestamp);
  sheet.getRange(row, headerMap['OPD Location - OPD']).setValue('OPD Unit');
  sheet.getRange(row, headerMap['Number of OPD Visits - OPD']).setValue(formData.numberOfOPDVisit);
  sheet.getRange(row, headerMap['Number of Nebulizations - OPD']).setValue(formData.numberOfNebulizationOpd);
  sheet.getRange(row, headerMap['Number of Injections - OPD']).setValue(formData.numberOfInjections);
  sheet.getRange(row, headerMap['Number of Dressings - OPD']).setValue(formData.numberOfDressing);
  sheet.getRange(row, headerMap['Number of Minor Surgeries - OPD']).setValue(formData.numberOfMinorSurgeries);
  sheet.getRange(row, headerMap['Number of OPD Umandawa - OPD']).setValue(formData.numberOfOPDUmandawa);
  sheet.getRange(row, headerMap['Number of Blood Drawings - OPD']).setValue(formData.numberOfBloodDrawing);
  sheet.getRange(row, headerMap['Number of Catheterizations - OPD']).setValue(formData.numberOfCatheterization);
}

function submitEndoscopy(formData) {
  var dateString = formData.date;
  var row = getOrCreateRow(dateString);
  var sheet = getSheet();
  var headerMap = getHeaderMap(sheet);

  var timestamp = new Date();
  sheet.getRange(row, headerMap['Endoscopy Timestamp - Endo']).setValue(timestamp);
  sheet.getRange(row, headerMap['Endoscopy Location - Endo']).setValue('Endoscopy Unit');
  sheet.getRange(row, headerMap['Total Number of Endoscopies - Endo']).setValue(formData.totalEndoscopies);
}

function submitBloodBank(formData) {
  var dateString = formData.date;
  var row = getOrCreateRow(dateString);
  var sheet = getSheet();
  var headerMap = getHeaderMap(sheet);

  var timestamp = new Date();
  sheet.getRange(row, headerMap['Blood Bank Timestamp - BB']).setValue(timestamp);
  sheet.getRange(row, headerMap['Blood Bank Location - BB']).setValue('Bloodbank');
  sheet.getRange(row, headerMap['No. of Blood Issuing - BB']).setValue(formData.bloodIssue);
  sheet.getRange(row, headerMap['No. Blood component Issuing - BB']).setValue(formData.componentIssuing);
  sheet.getRange(row, headerMap['No. of DT Requested - BB']).setValue(formData.dtRequested);
  sheet.getRange(row, headerMap['No. of cross Matching done - BB']).setValue(formData.crossMatching);
  sheet.getRange(row, headerMap['Blood issued to other hospitals - BB']).setValue(formData.issuedOtherHospitals);
  sheet.getRange(row, headerMap['Component Issued to other hospitals - BB']).setValue(formData.componentOtherHospitals);
  sheet.getRange(row, headerMap['Blood received from other  hospitals - BB']).setValue(formData.receivedOtherHospitals);
  sheet.getRange(row, headerMap['No of Component received from other hospitals - BB']).setValue(formData.componentReceivedOtherHospitals);
  sheet.getRange(row, headerMap['No of Component request - BB']).setValue(formData.componentRequest);
  sheet.getRange(row, headerMap['Mobile collection - BB']).setValue(formData.mobileCollection);
  sheet.getRange(row, headerMap['In house collection - BB']).setValue(formData.inHouseCollection);
  sheet.getRange(row, headerMap['Plt/ concentration received to other hospitals - BB']).setValue(formData.pltConcentrationIssued);
  sheet.getRange(row, headerMap['Total blood stock on (Pint) - BB']).setValue(formData.totalStock);
}

function submitHematology(formData) {
  var dateString = formData.date;
  var row = getOrCreateRow(dateString);
  var sheet = getSheet();
  var headerMap = getHeaderMap(sheet);

  var timestamp = new Date();
  sheet.getRange(row, headerMap['Hematology Timestamp - HAEM']).setValue(timestamp);
  sheet.getRange(row, headerMap['Hematology Location - HAEM']).setValue('Haematology Unit');
  sheet.getRange(row, headerMap['No. Of Admission - HAEM']).setValue(formData.admissions);
  sheet.getRange(row, headerMap['No. Of Clinic Pt.s - HAEM']).setValue(formData.clinicPts);
  sheet.getRange(row, headerMap['No. Phlebotomy - HAEM']).setValue(formData.phlebotomy);
  sheet.getRange(row, headerMap['Platelet - HAEM']).setValue(formData.platelet);
  sheet.getRange(row, headerMap['No of same day discharges - HAEM']).setValue(formData.sameDayDischarges);
  sheet.getRange(row, headerMap['No. Of Discharges - HAEM']).setValue(formData.discharges);
  sheet.getRange(row, headerMap['No. Of Referrals - HAEM']).setValue(formData.referrals);
  sheet.getRange(row, headerMap['No. of Bone marrow Biopsy - HAEM']).setValue(formData.boneMarrow);
  sheet.getRange(row, headerMap['SC injection (erythropoietin) - HAEM']).setValue(formData.scInjection);
}

function submitBloodDrawingCenter(formData) {
  var dateString = formData.date;
  var row = getOrCreateRow(dateString);
  var sheet = getSheet();
  var headerMap = getHeaderMap(sheet);

  var timestamp = new Date();
  sheet.getRange(row, headerMap['Blood Drawing Center Timestamp - BLEED']).setValue(timestamp);
  sheet.getRange(row, headerMap['Blood Drawing Center Location - BLEED']).setValue('Blood Drawing Centre');
  sheet.getRange(row, headerMap['Room No 38 - BLEED']).setValue(formData.room38);
  sheet.getRange(row, headerMap['Room No 23 - BLEED']).setValue(formData.room23);
  sheet.getRange(row, headerMap['Room No 04 - BLEED']).setValue(formData.room04);
  sheet.getRange(row, headerMap['Hematology - BLEED']).setValue(formData.hematology);
  sheet.getRange(row, headerMap['Total Numbers - BLEED']).setValue(formData.totalNumbers);
}

function submitInfectiousDiseases(formData) {
  var dateString = formData.date;
  var row = getOrCreateRow(dateString);
  var sheet = getSheet();
  var headerMap = getHeaderMap(sheet);

  var timestamp = new Date();
  sheet.getRange(row, headerMap['Infectious Diseases Timestamp - INFEC']).setValue(timestamp);
  sheet.getRange(row, headerMap['Infectious Diseases Location - INFEC']).setValue('Infection Control Unit');
  sheet.getRange(row, headerMap['Dengue - INFEC']).setValue(formData.dengue);
  sheet.getRange(row, headerMap['Leptospirosis - INFEC']).setValue(formData.leptospirosis);
  sheet.getRange(row, headerMap['H1N1 - INFEC']).setValue(formData.h1n1);
  sheet.getRange(row, headerMap['Influenza B - INFEC']).setValue(formData.influenzaB);
  sheet.getRange(row, headerMap['Covid 19 - INFEC']).setValue(formData.covid19);
  sheet.getRange(row, headerMap['Tetanus - INFEC']).setValue(formData.tetanus);
  sheet.getRange(row, headerMap['Chichenpox - INFEC']).setValue(formData.chickenpox);
  sheet.getRange(row, headerMap['Total  infections - INFEC']).setValue(formData.totalInfections);

}

function submitPainManagement(formData) {
  var dateString = formData.date;
  var row = getOrCreateRow(dateString);
  var sheet = getSheet();
  var headerMap = getHeaderMap(sheet);

  var timestamp = new Date();
  sheet.getRange(row, headerMap['Pain Management Timestamp - PMU']).setValue(timestamp);
  sheet.getRange(row, headerMap['Pain Management Location - PMU']).setValue('Pain Management Unit');
  sheet.getRange(row, headerMap['Post Operative Pain Assessment - PMU']).setValue(formData.postOpPainAssessment);
  sheet.getRange(row, headerMap['Monitoring of Epidural Patients - PMU']).setValue(formData.epiduralMonitoring);
  sheet.getRange(row, headerMap['Clinic Follow Up Patients - PMU']).setValue(formData.clinicFollowUp);
  sheet.getRange(row, headerMap['Maternal Patients - Back Massage - PMU']).setValue(formData.maternalBackMassage);
  sheet.getRange(row, headerMap['Maternal Patients - Health Education - PMU']).setValue(formData.maternalHealthEducation);
  sheet.getRange(row, headerMap['Non-Pharmacological Therapy - PMU']).setValue(formData.nonPharmaTherapy);
  sheet.getRange(row, headerMap['Total Number of Patients - PMU']).setValue(formData.totalPts);
}

function submitChemoTherapy(formData) {
  var dateString = formData.date;
  var row = getOrCreateRow(dateString);
  var sheet = getSheet();
  var headerMap = getHeaderMap(sheet);

  var timestamp = new Date();
  sheet.getRange(row, headerMap['Chemo Therapy Timestamp - CHEMO']).setValue(timestamp);
  sheet.getRange(row, headerMap['Chemo Therapy Location - CHEMO']).setValue('Chemotherapy Day Unit');
  sheet.getRange(row, headerMap['Number of Chemotherapy Treatments - CHEMO']).setValue(formData.chemoTherapy);
  sheet.getRange(row, headerMap['Number of Zoledronic Acid Treatments - CHEMO']).setValue(formData.zoledronicAcid);
  sheet.getRange(row, headerMap['Number of Filgrastim Treatments - CHEMO']).setValue(formData.filgrastimTreatment);
  sheet.getRange(row, headerMap['Number of Postponed Patients - CHEMO']).setValue(formData.postponedPts);
  sheet.getRange(row, headerMap['Total Number - CHEMO']).setValue(formData.total);
}

function submitBreastFeeding(formData) {
  var dateString = formData.date;
  var row = getOrCreateRow(dateString);
  var sheet = getSheet();
  var headerMap = getHeaderMap(sheet);

  var timestamp = new Date();
  sheet.getRange(row, headerMap['Breastfeeding Centre Timestamp - BF']).setValue(timestamp);
  sheet.getRange(row, headerMap['Breastfeeding Centre Location - BF']).setValue('Breastfeeding Day Centre');
  sheet.getRange(row, headerMap['From Out of Hospital Referrals - BF']).setValue(formData.outOfHospitalReferralsBf);
  sheet.getRange(row, headerMap['From Inward Referrals - BF']).setValue(formData.inwardReferralsBf);
  sheet.getRange(row, headerMap['Antenatal Mothers - BF']).setValue(formData.antenatalMothersBf);
  sheet.getRange(row, headerMap['Postnatal Mothers - BF']).setValue(formData.postnatalMothersBf);
  sheet.getRange(row, headerMap['Post & Antenatal mothers - BF']).setValue(formData.postAndAntenatalMothersBf);
  sheet.getRange(row, headerMap['From Discharge - BF']).setValue(formData.dischargeReferralsBf);
  sheet.getRange(row, headerMap['From Clinic/OPD Referrals - BF']).setValue(formData.opdReferralsBf);
  sheet.getRange(row, headerMap['Total - BF']).setValue(formData.totalBf);
}
function submitMaternalCensus(formData) {
  var dateString = formData.date;
  var row = getOrCreateRow(dateString);
  var sheet = getSheet();
  var headerMap = getHeaderMap(sheet);

  var timestamp = new Date();
  sheet.getRange(row, headerMap['Maternal Census Timestamp - MATERNAL']).setValue(timestamp);
  sheet.getRange(row, headerMap['Maternal Census Location - MATERNAL']).setValue('Maternal Complex');
  sheet.getRange(row, headerMap['Normal Vaginal Deliveries (NVDs) - MATERNAL']).setValue(formData.nvds);
  sheet.getRange(row, headerMap['Forceps Delivery - MATERNAL']).setValue(formData.forcepsDelivery);
  sheet.getRange(row, headerMap['Vacuum Delivery - MATERNAL']).setValue(formData.vacuumDelivery);
  sheet.getRange(row, headerMap['LSCS - MATERNAL']).setValue(formData.lscs);
  sheet.getRange(row, headerMap['Total Deliveries - MATERNAL']).setValue(formData.totalDeliveries);
  sheet.getRange(row, headerMap['Twin Deliveries - MATERNAL']).setValue(formData.twinDelivery);
  sheet.getRange(row, headerMap['Triplet Deliveries - MATERNAL']).setValue(formData.tripletDelivery);
  sheet.getRange(row, headerMap['IUD - MATERNAL']).setValue(formData.iud);
  sheet.getRange(row, headerMap['Still Births - MATERNAL']).setValue(formData.stillBirth);
  sheet.getRange(row, headerMap['Neonatal Deaths - MATERNAL']).setValue(formData.neonatalDeath);
  sheet.getRange(row, headerMap['Maternal Deaths - MATERNAL']).setValue(formData.maternalDeath);
}
function submitDialysisData(formData) {
  var dateString = formData.date;
  var row = getOrCreateRow(dateString);
  var sheet = getSheet();
  var headerMap = getHeaderMap(sheet);

  var timestamp = new Date();
  sheet.getRange(row, headerMap['Dialysis Unit Timestamp - DU']).setValue(timestamp);
  sheet.getRange(row, headerMap['Dialysis Unit Location - DU']).setValue('Dialysis Unit');
  sheet.getRange(row, headerMap['Direct - DU']).setValue(formData.directDialysis);
  sheet.getRange(row, headerMap['In Ward Patient - DU']).setValue(formData.inWardPatient);
  sheet.getRange(row, headerMap['Total Dialysis - DU']).setValue(formData.totalDialysis);
  sheet.getRange(row, headerMap['Renal Biopsy - DU']).setValue(formData.renalBiopsy);
  sheet.getRange(row, headerMap['CAPD Insert - DU']).setValue(formData.capdInsert);
  sheet.getRange(row, headerMap['Permanent HD Catheter Insert - DU']).setValue(formData.permanentHDCatheterInsert);
  sheet.getRange(row, headerMap['Temporary HD Catheter Insert - DU']).setValue(formData.tempHdCatheterInsert);
  sheet.getRange(row, headerMap['Other - DU']).setValue(formData.otherDialysis);
}

function submitMidnightSummary(formData) {
  var dateString = formData.date;
  var row = getOrCreateRow(dateString);
  var sheet = getSheet();
  var headerMap = getHeaderMap(sheet);

  var timestamp = new Date();
  sheet.getRange(row, headerMap['Midnight Summary Timestamp - AE']).setValue(timestamp);
  sheet.getRange(row, headerMap['Midnight Summary Location - AE']).setValue('A&E');
  sheet.getRange(row, headerMap['Total Admissions - AE']).setValue(formData.admissionsAe);
  sheet.getRange(row, headerMap['Hospital Admissions - AE']).setValue(formData.hospitalAdmissions);
  sheet.getRange(row, headerMap['Total Daily In - AE']).setValue(formData.totalDailyIn);
  sheet.getRange(row, headerMap['Transfers to Other Wards - AE']).setValue(formData.transfersToOtherWards);
  sheet.getRange(row, headerMap['Transfers to Other Hospitals - AE']).setValue(formData.transfersToOtherHospitals);
  sheet.getRange(row, headerMap['Total Discharges - AE']).setValue(formData.dischargesAe);
  sheet.getRange(row, headerMap['Same Day Discharges - AE']).setValue(formData.sameDayDischargesAe);
  sheet.getRange(row, headerMap['Missings - AE']).setValue(formData.missings);
  sheet.getRange(row, headerMap['Lamas - AE']).setValue(formData.lamas);
  sheet.getRange(row, headerMap['Number of Deaths - AE']).setValue(formData.numberOfDeaths);
  sheet.getRange(row, headerMap['On Admission Deaths - AE']).setValue(formData.onAdmissionDeaths);
  sheet.getRange(row, headerMap['Total Daily Out - AE']).setValue(formData.totalDailyOut);
  sheet.getRange(row, headerMap['Previous Day Midnight Total - AE']).setValue(formData.previousDayMidnightTotal);
  sheet.getRange(row, headerMap['Today Midnight Total - AE']).setValue(formData.todayMidnightTotal);
  sheet.getRange(row, headerMap['ARV - AE']).setValue(formData.arv);
  sheet.getRange(row, headerMap['ARS - AE']).setValue(formData.ars);
}

function submitTheaterAbData(formData) {
  var dateString = formData.date;
  var row = getOrCreateRow(dateString);
  var sheet = getSheet();
  var headerMap = getHeaderMap(sheet);

  var timestamp = new Date();
  sheet.getRange(row, headerMap['AB Theater Data Timestamp']).setValue(timestamp);
  sheet.getRange(row, headerMap['AB Theater Name']).setValue('AB');
  sheet.getRange(row, headerMap['Date']).setValue(formData.date);
  sheet.getRange(row, headerMap['AB Major Surgeries']).setValue(formData.majorAb);
  sheet.getRange(row, headerMap['AB Minor Surgeries']).setValue(formData.minorAb);
  sheet.getRange(row, headerMap['AB Total Surgeries']).setValue(formData.totalAb);
  sheet.getRange(row, headerMap['AB General Anesthesia (GA)']).setValue(formData.gaAb);
  sheet.getRange(row, headerMap['AB Local Anesthesia (LA)']).setValue(formData.laAb);
  sheet.getRange(row, headerMap['AB Spinal Anesthesia']).setValue(formData.spinalAb);
  sheet.getRange(row, headerMap['AB Epidural Anesthesia']).setValue(formData.epiduralAb);
  sheet.getRange(row, headerMap['AB Topical Anesthesia']).setValue(formData.topicalAb);
  sheet.getRange(row, headerMap['AB Other Anesthesia']).setValue(formData.otherAb);
  sheet.getRange(row, headerMap['AB Total']).setValue(formData.totalNetAb);
}
function submitTheaterCdData(formData) {
  var dateString = formData.date;
  var row = getOrCreateRow(dateString);
  var sheet = getSheet();
  var headerMap = getHeaderMap(sheet);

  var timestamp = new Date();
  sheet.getRange(row, headerMap['CD Theater Data Timestamp']).setValue(timestamp);
  sheet.getRange(row, headerMap['CD Theater Name']).setValue('CD');
  sheet.getRange(row, headerMap['Date']).setValue(formData.date);
  sheet.getRange(row, headerMap['CD Major Surgeries']).setValue(formData.majorCd);
  sheet.getRange(row, headerMap['CD Minor Surgeries']).setValue(formData.minorCd);
  sheet.getRange(row, headerMap['CD Total Surgeries']).setValue(formData.totalCd);
  sheet.getRange(row, headerMap['CD General Anesthesia (GA)']).setValue(formData.gaCd);
  sheet.getRange(row, headerMap['CD Local Anesthesia (LA)']).setValue(formData.laCd);
  sheet.getRange(row, headerMap['CD Spinal Anesthesia']).setValue(formData.spinalCd);
  sheet.getRange(row, headerMap['CD Epidural Anesthesia']).setValue(formData.epiduralCd);
  sheet.getRange(row, headerMap['CD Topical Anesthesia']).setValue(formData.topicalCd);
  sheet.getRange(row, headerMap['CD Other Anesthesia']).setValue(formData.otherCd);
  sheet.getRange(row, headerMap['CD Total']).setValue(formData.totalNetCd);
}
function submitTheaterEfData(formData) {
  var dateString = formData.date;
  var row = getOrCreateRow(dateString);
  var sheet = getSheet();
  var headerMap = getHeaderMap(sheet);

  var timestamp = new Date();
  sheet.getRange(row, headerMap['EF Theater Data Timestamp']).setValue(timestamp);
  sheet.getRange(row, headerMap['EF Theater Name']).setValue('EF');
  sheet.getRange(row, headerMap['Date']).setValue(formData.date);
  sheet.getRange(row, headerMap['EF Major Surgeries']).setValue(formData.majorEf);
  sheet.getRange(row, headerMap['EF Minor Surgeries']).setValue(formData.minorEf);
  sheet.getRange(row, headerMap['EF Total Surgeries']).setValue(formData.totalEf);
  sheet.getRange(row, headerMap['EF General Anesthesia (GA)']).setValue(formData.gaEf);
  sheet.getRange(row, headerMap['EF Local Anesthesia (LA)']).setValue(formData.laEf);
  sheet.getRange(row, headerMap['EF Spinal Anesthesia']).setValue(formData.spinalEf);
  sheet.getRange(row, headerMap['EF Epidural Anesthesia']).setValue(formData.epiduralEf);
  sheet.getRange(row, headerMap['EF Topical Anesthesia']).setValue(formData.topicalEf);
  sheet.getRange(row, headerMap['EF Other Anesthesia']).setValue(formData.otherEf);
  sheet.getRange(row, headerMap['EF Total']).setValue(formData.totalNetEf);
}
function submitTheaterGhData(formData) {
  var dateString = formData.date;
  var row = getOrCreateRow(dateString);
  var sheet = getSheet();
  var headerMap = getHeaderMap(sheet);

  var timestamp = new Date();
  sheet.getRange(row, headerMap['GH Theater Data Timestamp']).setValue(timestamp);
  sheet.getRange(row, headerMap['GH Theater Name']).setValue('GH');
  sheet.getRange(row, headerMap['Date']).setValue(formData.date);
  sheet.getRange(row, headerMap['GH Major Surgeries']).setValue(formData.majorGh);
  sheet.getRange(row, headerMap['GH Minor Surgeries']).setValue(formData.minorGh);
  sheet.getRange(row, headerMap['GH Total Surgeries']).setValue(formData.totalGh);
  sheet.getRange(row, headerMap['GH General Anesthesia (GA)']).setValue(formData.gaGh);
  sheet.getRange(row, headerMap['GH Local Anesthesia (LA)']).setValue(formData.laGh);
  sheet.getRange(row, headerMap['GH Spinal Anesthesia']).setValue(formData.spinalGh);
  sheet.getRange(row, headerMap['GH Epidural Anesthesia']).setValue(formData.epiduralGh);
  sheet.getRange(row, headerMap['GH Topical Anesthesia']).setValue(formData.topicalGh);
  sheet.getRange(row, headerMap['GH Other Anesthesia']).setValue(formData.otherGh);
  sheet.getRange(row, headerMap['GH Total']).setValue(formData.totalNetGh);
}
function submitTheaterMData(formData) {
  var dateString = formData.date;
  var row = getOrCreateRow(dateString);
  var sheet = getSheet();
  var headerMap = getHeaderMap(sheet);

  var timestamp = new Date();
  sheet.getRange(row, headerMap['M Theater Data Timestamp']).setValue(timestamp);
  sheet.getRange(row, headerMap['M Theater Name']).setValue('M');
  sheet.getRange(row, headerMap['Date']).setValue(formData.date);
  sheet.getRange(row, headerMap['M Major Surgeries']).setValue(formData.majorM);
  sheet.getRange(row, headerMap['M Minor Surgeries']).setValue(formData.minorM);
  sheet.getRange(row, headerMap['M Total Surgeries']).setValue(formData.totalM);
  sheet.getRange(row, headerMap['M General Anesthesia (GA)']).setValue(formData.gaM);
  sheet.getRange(row, headerMap['M Local Anesthesia (LA)']).setValue(formData.laM);
  sheet.getRange(row, headerMap['M Spinal Anesthesia']).setValue(formData.spinalM);
  sheet.getRange(row, headerMap['M Epidural Anesthesia']).setValue(formData.epiduralM);
  sheet.getRange(row, headerMap['M Topical Anesthesia']).setValue(formData.topicalM);
  sheet.getRange(row, headerMap['M Other Anesthesia']).setValue(formData.otherM);
  sheet.getRange(row, headerMap['M Total']).setValue(formData.totalNetM);
}
function submitTheaterEyeData(formData) {
  var dateString = formData.date;
  var row = getOrCreateRow(dateString);
  var sheet = getSheet();
  var headerMap = getHeaderMap(sheet);

  var timestamp = new Date();
  sheet.getRange(row, headerMap['Eye Theater Data Timestamp']).setValue(timestamp);
  sheet.getRange(row, headerMap['Eye Theater Name']).setValue('Eye');
  sheet.getRange(row, headerMap['Date']).setValue(formData.date);
  sheet.getRange(row, headerMap['Eye Major Surgeries']).setValue(formData.majorI);
  sheet.getRange(row, headerMap['Eye Minor Surgeries']).setValue(formData.minorI);
  sheet.getRange(row, headerMap['Eye Total Surgeries']).setValue(formData.totalI);
  sheet.getRange(row, headerMap['Eye General Anesthesia (GA)']).setValue(formData.gaI);
  sheet.getRange(row, headerMap['Eye Local Anesthesia (LA)']).setValue(formData.laI);
  sheet.getRange(row, headerMap['Eye Spinal Anesthesia']).setValue(formData.spinalI);
  sheet.getRange(row, headerMap['Eye Epidural Anesthesia']).setValue(formData.epiduralI);
  sheet.getRange(row, headerMap['Eye Topical Anesthesia']).setValue(formData.topicalI);
  sheet.getRange(row, headerMap['Eye Other Anesthesia']).setValue(formData.otherI);
  sheet.getRange(row, headerMap['Eye Total']).setValue(formData.totalNetI);
}
/**
 * Returns submission status for a given date
 * @param {string} dateString in 'YYYY-MM-DD'
 * @return {{opdTimestamp:string, endoscopyTimestamp:string}}
 */
function getSubmissionStatus(dateString) {
  var sheet = getSheet();
  var headerMap = getHeaderMap(sheet);
  var row = getOrCreateRow(dateString);

  // Retrieving timestamps for each service
  var opdTs = sheet.getRange(row, headerMap['OPD Timestamp']).getValue();
  var endoTs = sheet.getRange(row, headerMap['Endoscopy Timestamp']).getValue();
  var bloodBankTs = sheet.getRange(row, headerMap['Blood Bank Timestamp']).getValue();
  var hematologyTs = sheet.getRange(row, headerMap['Hematology Timestamp']).getValue();
  var bloodDrawingCenterTs = sheet.getRange(row, headerMap['Blood Drawing Center Timestamp']).getValue();
  var infectiousDiseasesTs = sheet.getRange(row, headerMap['Infectious Diseases Timestamp']).getValue();
  var painManagementTs = sheet.getRange(row, headerMap['Pain Management Timestamp']).getValue();
  var chemoTherapyTs = sheet.getRange(row, headerMap['Chemo Therapy Timestamp']).getValue();
  var breastFeedingTs = sheet.getRange(row, headerMap['Breast Feeding Timestamp']).getValue();
  var maternalCensusTs = sheet.getRange(row, headerMap['Maternal Census Timestamp']).getValue();
  var dialysisTs = sheet.getRange(row, headerMap['Dialysis Timestamp']).getValue();
  var midnightSummaryTs = sheet.getRange(row, headerMap['Midnight Summary Timestamp']).getValue();
  var theaterAbDataTs = sheet.getRange(row, headerMap['AB Theater Data Timestamp']).getValue();
  var theaterCdDataTs = sheet.getRange(row, headerMap['CD Theater Data Timestamp']).getValue();
  var theaterEfDataTs = sheet.getRange(row, headerMap['EF Theater Data Timestamp']).getValue();
  var theaterGhDataTs = sheet.getRange(row, headerMap['GH Theater Data Timestamp']).getValue();
  var theaterMDataTs = sheet.getRange(row, headerMap['M Theater Data Timestamp']).getValue();
  var theaterEyeDataTs = sheet.getRange(row, headerMap['Eye Theater Data Timestamp']).getValue();

  return {
    opdTimestamp: opdTs ? opdTs.toString() : '',
    endoscopyTimestamp: endoTs ? endoTs.toString() : '',
    bloodBankTimestamp: bloodBankTs ? bloodBankTs.toString() : '',
    hematologyTimestamp: hematologyTs ? hematologyTs.toString() : '',
    bloodDrawingCenterTimestamp: bloodDrawingCenterTs ? bloodDrawingCenterTs.toString() : '',
    infectiousDiseasesTimestamp: infectiousDiseasesTs ? infectiousDiseasesTs.toString() : '',
    painManagementTimestamp: painManagementTs ? painManagementTs.toString() : '',
    chemoTherapyTimestamp: chemoTherapyTs ? chemoTherapyTs.toString() : '',
    breastFeedingTimestamp: breastFeedingTs ? breastFeedingTs.toString() : '',
    maternalCensusTimestamp: maternalCensusTs ? maternalCensusTs.toString() : '',
    dialysisTimestamp: dialysisTs ? dialysisTs.toString() : '',
    midnightSummaryTimestamp: midnightSummaryTs ? midnightSummaryTs.toString() : '',
    theaterAbDataTimestamp: theaterAbDataTs ? theaterAbDataTs.toString() : '',
    theaterCdDataTimestamp: theaterCdDataTs ? theaterCdDataTs.toString() : '',
    theaterEfDataTimestamp: theaterEfDataTs ? theaterEfDataTs.toString() : '',
    theaterGhDataTimestamp: theaterGhDataTs ? theaterGhDataTs.toString() : '',
    theaterMDataTimestamp: theaterMDataTs ? theaterMDataTs.toString() : '',
    theaterEyeDataTimestamp: theaterEyeDataTs ? theaterEyeDataTs.toString() : ''
  };
}


/**
 * Retrieves the Data sheet
 */
function getSheet() {
  return SpreadsheetApp.openById('1cWB2BNHvKXm9jIjTOFxpDREMSvHj5JCZ6r-PI4a6GeU').getSheetByName('Data');
}

/**
 * Builds a map of headerName -> column number
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet
 * @return {Object<string, number>}
 */
function getHeaderMap(sheet) {
  var headerRow = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  var map = {};
  for (var i = 0; i < headerRow.length; i++) {
    map[headerRow[i]] = i + 1;
  }
  return map;
}

/**
 * Finds the row for given date (YYYY-MM-DD); creates it if missing
 * @param {string} dateString
 * @return {number} row number
 */
function getOrCreateRow(dateString) {
  var sheet = getSheet();
  var headerMap = getHeaderMap(sheet);
  var dateCol = headerMap['Date'];
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) {
    sheet.appendRow([new Date(dateString)]);
    return 2;
  }
  var dates = sheet.getRange(2, dateCol, lastRow - 1, 1).getValues();
  for (var i = 0; i < dates.length; i++) {
    var cell = dates[i][0];
    var ssDate = '';
    if (cell instanceof Date) {
      ssDate = Utilities.formatDate(cell, Session.getScriptTimeZone(), 'yyyy-MM-dd');
    } else {
      ssDate = cell.toString();
    }
    if (ssDate === dateString) {
      return i + 2;
    }
  }
  // not found → append new row with date
  sheet.appendRow([new Date(dateString)]);
  return sheet.getLastRow();
}
