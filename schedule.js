
// import moment from './moment.min';
// import es from './es.min';
var monthsShortDot="ene._feb._mar._abr._may._jun._jul._ago._sep._oct._nov._dic.".split("_"),monthsShort="ene_feb_mar_abr_may_jun_jul_ago_sep_oct_nov_dic".split("_");moment.defineLocale("es",{months:"enero_febrero_marzo_abril_mayo_junio_julio_agosto_septiembre_octubre_noviembre_diciembre".split("_"),monthsShort:function(e,s){return e?/-MMM-/.test(s)?monthsShort[e.month()]:monthsShortDot[e.month()]:monthsShortDot},monthsParseExact:!0,weekdays:"domingo_lunes_martes_miércoles_jueves_viernes_sábado".split("_"),weekdaysShort:"dom._lun._mar._mié._jue._vie._sáb.".split("_"),weekdaysMin:"do_lu_ma_mi_ju_vi_sá".split("_"),weekdaysParseExact:!0,longDateFormat:{LT:"H:mm",LTS:"H:mm:ss",L:"DD/MM/YYYY",LL:"D [de] MMMM [de] YYYY",LLL:"D [de] MMMM [de] YYYY H:mm",LLLL:"dddd, D [de] MMMM [de] YYYY H:mm"},calendar:{sameDay:function(){return"[hoy a la"+(1!==this.hours()?"s":"")+"] LT"},nextDay:function(){return"[mañana a la"+(1!==this.hours()?"s":"")+"] LT"},nextWeek:function(){return"dddd [a la"+(1!==this.hours()?"s":"")+"] LT"},lastDay:function(){return"[ayer a la"+(1!==this.hours()?"s":"")+"] LT"},lastWeek:function(){return"[el] dddd [pasado a la"+(1!==this.hours()?"s":"")+"] LT"},sameElse:"L"},relativeTime:{future:"en %s",past:"hace %s",s:"unos segundos",m:"un minuto",mm:"%d minutos",h:"una hora",hh:"%d horas",d:"un día",dd:"%d días",M:"un mes",MM:"%d meses",y:"un año",yy:"%d años"},dayOfMonthOrdinalParse:/\d{1,2}º/,ordinal:"%dº",week:{dow:1,doy:4}});

moment.locale('es');
moment.updateLocale('es', {
	monthsShort: moment.monthsShort('-MMM-')
});
// {1} = Date created
// {2} = Date start with time
// {3} = Date end with timestamp
// {4} = Location
// {5} = UNTIL (datetimestamp)
// {6} = Day frequency (MO,TU,WE,TH,FR) and any combination thereof
// {7} = Name of the event
if (!String.prototype.format) {
	String.prototype.format = function() {
		var args = arguments;
		return this.replace(/{(\d+)}/g, function(match, number) {
			return typeof args[number] != 'undefined' ? args[number] : match;
		});
	};
}
var dayseng = ["SU", "MO", "TU", "WE", "TH", "FR", "SA"];

var data = "BEGIN:VCALENDAR\nPRODID:-//Microsoft Corporation//Outlook 16.0 MIMEDIR//EN\nVERSION:2.0\nMETHOD:PUBLISH\nBEGIN:VTIMEZONE\nTZID:Central Standard Time\nBEGIN:STANDARD\nDTSTART:16011104T020000\nRRULE:FREQ=YEARLY;BYDAY=1SU;BYMONTH=11\nTZOFFSETFROM:-0500\nTZOFFSETTO:-0600\nEND:STANDARD\nBEGIN:DAYLIGHT\nDTSTART:16010311T020000\nRRULE:FREQ=YEARLY;BYDAY=2SU;BYMONTH=3\nTZOFFSETFROM:-0600\nTZOFFSETTO:-0500\nEND:DAYLIGHT\nEND:VTIMEZONE\n";
// var data = "BEGIN:VCALENDAR\nPRODID:-//Microsoft Corporation//Outlook 16.0 MIMEDIR//EN\nVERSION:2.0\nMETHOD:PUBLISH\nBEGIN:VTIMEZONE\nTZID:Pacific Standard Time\nBEGIN:STANDARD\nDTSTART:16011104T020000\nRRULE:FREQ=YEARLY;BYDAY=1SU;BYMONTH=11\nTZOFFSETFROM:-0700\nTZOFFSETTO:-0800\nEND:STANDARD\nBEGIN:DAYLIGHT\nDTSTART:16010311T020000\nRRULE:FREQ=YEARLY;BYDAY=2SU;BYMONTH=3\nTZOFFSETFROM:-0800\nTZOFFSETTO:-0700\nEND:DAYLIGHT\nEND:VTIMEZONE";

var classTmpl = 'BEGIN:VEVENT\nCLASS:PUBLIC\nCREATED:{0}\nUID:{0}-123401@itesm.mx\nDESCRIPTION: \\n\nDTSTART;TZID="Central Standard Time":{1}\nDTEND;TZID="Central Standard Time":{2}\nDTSTAMP:{0}\nLOCATION:{3}\nRRULE:FREQ=WEEKLY;UNTIL={4};BYDAY={5}\nSEQUENCE:0\nSUMMARY;LANGUAGE=en-us:{6}\nTRANSP:OPAQUE\nEND:VEVENT\n';



var msgBody = document.getElementById('Item.MessageNormalizedBody');

var t1 = msgBody.querySelectorAll('center')[2];
var tableDesc = t1.querySelectorAll('table')[1];
var tbody = tableDesc.children[0];
var trs = tbody.children;
var subjectRows = tbody.querySelectorAll('tr[bgcolor="#9bbad6"]');

var subjects = [];
for(var i = 0; i < subjectRows.length; i++) {
	subjects.push({
		name: subjectRows[i].children[1].querySelector('code').textContent.trim(),
		details: getDetails(subjectRows[i])
	});
}

subjects.forEach(function(subject) {

	// {1}
	var datets = new Date().toISOString().replace(/-/g, '');
	datets = datets.substring(0, datets.indexOf('.')).replace(/:/g, '') + 'Z';

	subject.details.forEach(function(evento) {
		var dates = evento.fechas.split('al');

		var horas = evento.horario.split(' a ');

		// {2}
		var startDate = dates[0].trim().toLowerCase();
		var endTimeDate = parseDate(startDate + " " + horas[1]);
		startDate += " " + horas[0];

		// {3}
		var endDate = dates[1].trim().toLowerCase() + " " + horas[1];

		startDate = parseDate(startDate);
		endDate = parseDate(endDate);
		// {4}
		var location = evento.salon;

		// {6}
		var weekdays = "";
		for(var i = 0; i < evento.dias.length; i++) {
			if(evento.dias[i].trim() !== '') {
				if(weekdays !== "") {
					weekdays += ',';
				}
				weekdays += dayseng[i];
			}
		}
		data += classTmpl.format(datets, startDate, endTimeDate, location, endDate, weekdays, subject.name);
	});
});
function parseDate(str) {
	return replaceChars(moment(str, 'D [de] MMM YYYY H:mm', 'es', true).format());
}
function replaceChars(str){
	var cutoff = str.lastIndexOf('-');
	var ns = str.substring(0, cutoff);
	return ns.replace(/-/g, '').replace(/:/g, '');
}

function getDetails(elem) {
	var el = elem.nextElementSibling;
	var details = [];
	while(el && el.getAttribute('bgcolor') !== '#9bbad6') {
		if(el.getAttribute('bgcolor') === '#EEEEEE') {
			var fechas = el.children[1].textContent.trim();
			var dias = el.children[2].textContent;
			var horario = el.children[3].textContent.trim();
			var tex4 = el.children[4].textContent.trim();
			var tex5 = el.children[5].textContent.trim();
			var salon;
			if(tex4.substr(0,2) == 'A-')
				tex4 = tex4.substr(2);
			salon = tex4 == tex5 ? tex5 : tex4 + tex5;
			details.push({
				fechas: fechas,
				salon: salon,
				horario: horario,
				dias: dias
			});
		}
		el = el.nextElementSibling;
	}
	return details;
}

data += "END:VCALENDAR";


function downloadICS() {
	var MIME_TYPE = "text/calendar";
	var blob = new Blob([data], {type: MIME_TYPE});
	window.location.href = window.URL.createObjectURL(blob);
}