<!--# -*- coding: utf-8 -*- -->
window.onload = function () {
	var chart1 = new CanvasJS.Chart("entries_chart_1", {
		title:{
			text: "回期別 参加申込み人数（回期順）"
		},
		data: [
			{
				type: 'column',
				dataPoints: dataPlot1
			}
		],
		exportFileName: "EntriesChart1",
		exportEnabled: true,
		axisX:{
			title: '卒業回期'
		},
		axisY:{
			title: '申込み人数',
      suffix: "人"
    },
	});

	var chart2 = new CanvasJS.Chart("entries_chart_2", {
		title:{
			text: "回期別 参加申込み人数（申し込み人数順）"
		},
		data: [
			{
				type: 'column',
				dataPoints: dataPlot2
			}
		],
		exportFileName: "EntriesChart2",
		exportEnabled: true,
		axisX:{
			title: '卒業回期',
      interval: 1,
			labelFontSize: 12,
      // labelAngle: -45
		},
		axisY:{
			title: '申込み人数',
      suffix: "人"
    },
	});

	chart1.render();
	chart2.render();
}
