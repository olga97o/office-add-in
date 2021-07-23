import {Component} from '@angular/core';
import InsertShiftDirection = Excel.InsertShiftDirection;
import EventHandlerResult = OfficeExtension.EventHandlerResult;
import {max} from "rxjs/operators";

@Component({
    selector: 'app-root',
    templateUrl: './app.component.html',
    styles: []
})

export class AppComponent {
    title = 'test-app1';

    fillRandomNum() {
        Excel.run(context => {
            const range = context.workbook.getSelectedRange();
            range.load(["address", "columnCount", "rowCount", "values"]);
            return context.sync().then(() => {
                console.log('range', range);
                let table = Array.from(range.values, x => x.map(cell => (Math.random() * (100 - 10) + 10)));
                range.set({
                    values: table,
                    format: {
                        fill: {
                            color: "#4472C4",
                            pattern: "Checker"
                        },
                        font: {
                            name: "Verdana",
                            color: "white"
                        },
                    }
                })
            })
        })
    }

    showAddressRange() {
        Excel.run(context => {
            const range = context.workbook.worksheets.getActiveWorksheet();
            range.onSelectionChanged.add((event) => {
                return Excel.run(context => {
                    let showDiv = document.getElementById("showAddressDiv");
                    console.log('in')
                    console.log("The selected range has changed to: " + event.address);
                    // @ts-ignore
                    showDiv.innerText = `Your select address: ${event.address}`;
                    return context.sync();
                });
            });
            console.log('out')
            return context.sync();
        });
    }

    splitForSquares() {
        let parts = Bounds.splitBounds(new Bounds(16, 2, 50, 15), 10, 0);
        //parts.forEach(p => console.log(p.toString()));
        //console.log("parts", parts);
        parts.forEach(el => {
            Excel.run(context => {
                const sheetName = "Sheet1";
                let rangeAddress: string;
                let range: Excel.Range;
                let randomColor: string;
                console.log('el', el)

                rangeAddress = this.fromNumToChar(el.x) + el.y + ':' + this.fromNumToChar(el.x + el.columnCount) + (el.rowCount);//'A1:C4'
                console.log('rangeAddress', rangeAddress);
                range = context.workbook.worksheets.getItem(sheetName).getRange(rangeAddress);
                range.load(["values", "address", "columnCount", "rowCount", "cellCount"]);
                randomColor = Math.floor(Math.random() * 16777215).toString(16);

                return context.sync().then(() => {
                    range.set({
                        format: {
                            fill: {
                                color: randomColor,
                                pattern: "Checker"
                            },
                            font: {
                                name: "Verdana",
                                color: "white"
                            },
                        }
                    })
                })

            })
        })
    }

    addPlusOne() {
        Excel.run(context => {
            const range = context.workbook.getSelectedRange();
            range.load(["values"]);
            return context.sync().then(() => {
                //console.log('range', range);
                let table = Array.from(range.values, x => x.map(cell => cell + 1));
                range.set({
                    values: table,
                    format: {
                        fill: {
                            color: "yellow",
                            pattern: "Checker"
                        },
                        font: {
                            name: "Verdana",
                            color: "black"
                        },
                    }
                })
            })
        })
        /*console.log('beforeCallPromise')
        this.testPromise().then(() => console.log('run promise'))*/
    }

    fromNumToChar(num: number) {
        let letterAddress;
        let secondLetter, firstLetter: string;
        if (num >= 26) {
            if (num % 26) {
                firstLetter = String.fromCharCode(64 + (num - (num % 26)) / 26);
                secondLetter = String.fromCharCode(64 + (num % 26));
            } else {
                firstLetter = String.fromCharCode(64 + (num - (num % 26)) / 26 - 1);
                secondLetter = String.fromCharCode(64 + (num % 26) + 26);
            }
            letterAddress = firstLetter + secondLetter;
        } else {
            letterAddress = String.fromCharCode(64 + num);
        }
        return letterAddress;
    }
}

export class Bounds {
    x: number = 0;
    y: number = 0;
    rowCount: number = 0;
    columnCount: number = 0;

    constructor(x: number, y: number, rowCount: number, columnCount: number) {
        this.x = x;
        this.y = y;
        this.rowCount = rowCount;
        this.columnCount = columnCount;
    }

    toString() {
        return `x: ${this.x} y:${this.y} rowCount: ${this.rowCount} columnCount: ${this.columnCount}`
    }

    static splitBounds(bounds: Bounds, maxCellsCount: number, maxColumnCount: number): Bounds[] {
        let counterSquaresTail: number = (bounds.rowCount * bounds.columnCount) % maxCellsCount;//3400
        let counterSquares: number = ((bounds.rowCount * bounds.columnCount) - counterSquaresTail) / maxCellsCount;//75
        let arr: Bounds[] = [];
        let lastEndPoint: number = 0;
        let endPoint: number = 0;
        let startPointX: number = bounds.x;
        let startPointY: number = bounds.y;
        let counter: number;

        for (let i: number = 1; i <= counterSquares; i++) {
            if (maxCellsCount <= bounds.columnCount || maxCellsCount <= bounds.rowCount) {
                if (bounds.rowCount < bounds.columnCount) {
                    let rowCounter: number = bounds.columnCount / maxCellsCount;
                    let columnCounter: number = maxCellsCount / rowCounter;
                    endPoint += columnCounter;
                    if (startPointY + rowCounter >= bounds.rowCount) {
                        startPointY = bounds.y;
                        startPointX += bounds.columnCount;
                    }
                    arr.push(new Bounds(startPointX, startPointY, rowCounter, columnCounter));
                    lastEndPoint = endPoint;
                    startPointX += columnCounter;
                } else {
                    let columnCounter: number = bounds.rowCount / maxCellsCount;//5
                    let rowCounter: number = maxCellsCount / columnCounter;//2
                    endPoint += rowCounter;
                    if (startPointX + columnCounter >= bounds.columnCount) {
                        startPointX = bounds.x;
                        startPointY += bounds.rowCount;
                    }

                    arr.push(new Bounds(startPointX, startPointY, rowCounter, columnCounter));
                    lastEndPoint = endPoint;
                    startPointY += rowCounter;
                }
            } else {
                if (bounds.rowCount < bounds.columnCount) {
                    counter = maxCellsCount / bounds.rowCount;
                    endPoint += counter;
                    arr.push(new Bounds(startPointX, startPointY, bounds.rowCount, counter));
                    lastEndPoint = endPoint;
                    startPointX += counter;
                } else {
                    counter = maxCellsCount / bounds.columnCount;//2
                    endPoint += counter;
                    arr.push(new Bounds(startPointX, startPointY, counter, bounds.columnCount));
                    lastEndPoint = endPoint;
                    startPointY += counter;
                }
            }
        }

        if (counterSquaresTail) {
            if (maxCellsCount <= bounds.columnCount || maxCellsCount <= bounds.rowCount) {
                /*if (bounds.rowCount < bounds.columnCount) {
                    let counterTail: number = counterSquaresTail / bounds.rowCount;//17
                    let startPointTail: number = lastEndPoint + 1;
                    arr.push(new Bounds(startPointTail, startPointY, bounds.rowCount, counterTail));
                    console.log('myArr', arr);
                } else {
                    let counterTail: number = counterSquaresTail / bounds.columnCount;//17
                    let startPointTail: number = lastEndPoint + 1;
                    arr.push(new Bounds(startPointX, startPointTail, counterTail, bounds.columnCount));
                }*/
            } else {
                if (bounds.rowCount < bounds.columnCount) {
                    let counterTail: number = counterSquaresTail / bounds.rowCount;//17
                    let startPointTail: number = lastEndPoint + 1;
                    arr.push(new Bounds(startPointTail, startPointY, bounds.rowCount, counterTail));
                    console.log('myArr', arr);
                } else {
                    let counterTail: number = counterSquaresTail / bounds.columnCount;//17
                    let startPointTail: number = lastEndPoint + 1;
                    arr.push(new Bounds(startPointX, startPointTail, counterTail, bounds.columnCount));
                }
            }
        }
        return arr;
    }
}
