/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global console, document, Excel, Office */

import QRCode from "qrcode";

// The initialize function must be run each time a new page is loaded
Office.initialize = () => {
  //	document.getElementById("app-body").style.display = "flex";
  //	document.getElementById("run").onclick = run;
    document.getElementById("getqrcode").onclick = getqrcode;
    document.getElementById("qrmake").onclick = qrmake;
 
  };
  

//	qrcode取得
async function getqrcode(){
	try {
		await Excel.run(async context => {
			//QRCode
    		let r = context.workbook.getSelectedRange();
			r.load("rowCount");
			await context.sync();

			let rowCount = r.rowCount;
			r = r.getAbsoluteResizedRange(rowCount, 2);

			let opts = {
				errorCorrectionLevel: 'H',
				type: 'image/png',
				//quality: 0.3,
				margin: 1,
				//color: {
				//	dark:"#010599FF",
				//	light:"#FFBF60FF"
				//}
			}

			r.load("items");
			await context.sync();

			for(let rowIndex=0; rowIndex<rowCount; rowIndex++){
				let cell = r.getCell(rowIndex,0);
				cell.load("text");
				await context.sync();
				let text = cell.text[0][0];
	
				let str = '';
				QRCode.toDataURL(text, opts, (err, url)=>{
					if (err) throw err
					str = url.replace('data:image/png;base64,','');
				});
				let shapes = context.workbook.worksheets.getActiveWorksheet().shapes;
				let shape = shapes.addImage(str);
				shape.lockAspectRatio = true;
				shape.placement = Excel.Placement.oneCell;
				cell = r.getCell(rowIndex,1);
				cell.load(['top','left']);
				await context.sync();
				shape.left = cell.left;
				shape.top = cell.top;
	
			}
/*
			let cell = r.getCell(0,0);
			cell.load("text");
			await context.sync();
			let text = cell.text[0][0];
			let codeinput = document.getElementById("qrcode_text") as HTMLInputElement;
			codeinput.value = text;
			
			qrmake_sub(text);

			let opts = {
				errorCorrectionLevel: 'H',
				type: 'image/png',
				//quality: 0.3,
				margin: 1,
				//color: {
				//	dark:"#010599FF",
				//	light:"#FFBF60FF"
				//}
			}
	
			let str = '';
			QRCode.toDataURL(text, opts, (err, url)=>{
				if (err) throw err
				str = url.replace('data:image/png;base64,','');
			});
			let shapes = context.workbook.worksheets.getActiveWorksheet().shapes;
			let shape = shapes.addImage(str);
			shape.lockAspectRatio = true;
			shape.placement = Excel.Placement.oneCell;
			cell = r.getCell(0,1);
			cell.load(['top','left']);
			await context.sync();
			shape.left = cell.left;
			shape.top = cell.top;
*/

		});
	} catch (error) {
		console.error(error);
	}
}
//	qrcode作成
async function qrmake(){
	try {
		await Excel.run(async context => {
			let codeinput = document.getElementById("qrcode_text") as HTMLInputElement;
			//console.log(codeinput.value);
			qrmake_sub(codeinput.value);
			//await context.sync();
		});
	} catch (error) {
		console.error(error);
	}
}

function qrmake_sub(code: string): boolean{
	let canvas = document.getElementById("qrcode") as HTMLCanvasElement;
	QRCode.toCanvas(canvas, code, (error)=>{
		if(error) {
			console.error(error);
			return false;
		}
	})
	return true;
}