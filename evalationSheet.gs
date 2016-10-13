function Person(name,sum,sumCell,average,aveCell,count){
  this.name = name;
  this.sum = sum;
  this.sumCell = sumCell;
  this.average = average;
  this.aveCell = aveCell;
  this.count = count;
}

function onOpen() {

  var numOfSheetsBefore = 10;
  var numOfSheetsAfter = 3;

  var NhatAnh = new Person("Nguyen Thi Nhat Anh", 0, 'B2', 0, 'B3', 0);
  var Hoang = new Person("Nguyen Tien Hoang", 0, 'D2', 0, 'D3', 0);
  var Thao = new Person("Nguyen Thi Phuong Thao", 0, 'F2', 0, 'F3', 0);
  var Huyen = new Person("Pham Thanh Huyen", 0, 'H2', 0, 'H3', 0);
  var Takuma = new Person("Takuma Hanaya", 0, 'J2', 0, 'J3', 0);

  var memberObjs = [NhatAnh, Hoang, Thao, Huyen, Takuma];

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  // 評価シート
  var evalationStr = "評価";
  var evalationSheet = ss.getSheetByName(evalationStr);
  // 全てのシート
  var sheets = ss.getSheets();
  sheets = sheets.reverse();

  // ***********集計する範囲のシート************
  sheets = sheets.slice(numOfSheetsBefore, sheets.length - numOfSheetsAfter);

  // 例{ 95: sheet（オブジェクト） }
  var sheetsObj = {};

  // 連想配列を作る
  sheetsObj = sheets.map(function(sheet, index) {
    return sheetsObj[sheet.getName().replace(/\//g, "")] = sheet;
  });

  // 評価シート内の、値を書き出す列の始まり
  var startRow = 5;
  for(var key in sheetsObj) {

    // 3行目全部チェック
    var authorNames = sheetsObj[key].getRange(3, 2, 1, 17).getValues()[0];
    authorNames.splice(9,7);
    authorNames.splice(1,7);
    authorNames.map(function(authorName) {
      if(authorName){
        Logger.log(authorName);
      }
    });

    var rowD = sheetsObj[key].getRange(16, 4, 17, 1).getValues();
    var rowL = sheetsObj[key].getRange(16, 12, 17, 1).getValues();
    var rowT = sheetsObj[key].getRange(16, 20, 17, 1).getValues();
    var rowAB = sheetsObj[key].getRange(16, 28, 17, 1).getValues();

    var Dscores = [];
    var Lscores = [];
    var Tscores = [];
    var ABscores = [];


    rowD.map(function(score) {
      if(score[0]) {
        Dscores.push(score[0]);
      }
    });
    rowL.map(function(score) {
      if(score[0]) {
        Lscores.push(score[0]);
      }
    });
    rowT.map(function(score) {
      if(score[0]) {
        Tscores.push(score[0]);
      }
    });
    rowAB.map(function(score) {
      if(score[0]) {
        ABscores.push(score[0]);
      }
    });

    authorNames.map(function(name, index) {

      var dailyScore = 0;
      var dailyScores = [];
      var average = 0;

      var calcAverage = function() {
        dailyScores.sort(function(a,b){
          if( a < b ) return -1;
          if( a > b ) return 1;
          return 0;
        });
        dailyScores.shift();
        dailyScores.pop();
        var sum = 0;
        dailyScores.map(function(k){
          sum += k;
        });
        return sum / dailyScores.length;
      };

      var echoAverageFrom = function(row, setRow) {
        row.map(callback);
        average = calcAverage();
        evalationSheet.getRange(startRow, setRow).setValue(average);
      };

      if(name == NhatAnh.name) {
        NhatAnh.count++;
        var callback = function(score) {
          NhatAnh.sum += score;
          dailyScores.push(score);
        };
        switch(index) {
          case 0:
            echoAverageFrom(Dscores, 2);
            break;
          case 1:
            echoAverageFrom(Lscores, 2);
            break;
          case 2:
            echoAverageFrom(Tscores, 2);
            break;
          case 3:
            echoAverageFrom(ABscores, 2);
            break;
          default:
            break;
        }
      }

      if(name == Hoang.name) {
        Hoang.count++;
        var callback = function(score) {
          Hoang.sum += score;
          dailyScores.push(score);
        };
        switch(index) {
          case 0:
            echoAverageFrom(Dscores, 4);
            break;
          case 1:
            echoAverageFrom(Lscores, 4);
            break;
          case 2:
            echoAverageFrom(Tscores, 4);
            break;
          case 3:
            echoAverageFrom(ABscores, 4);
            break;
          default:
            break;
        }
      }

      if(name == Thao.name) {
        Thao.count++;
        var callback = function(score) {
          Thao.sum += score;
          dailyScores.push(score);
        };
        switch(index) {
          case 0:
            echoAverageFrom(Dscores, 6);
            break;
          case 1:
            echoAverageFrom(Lscores, 6);
            break;
          case 2:
            echoAverageFrom(Tscores, 6);
            break;
          case 3:
            echoAverageFrom(ABscores, 6);
            break;
          default:
            break;
        }
      }

      if(name == Huyen.name) {
        Huyen.count++;
        var callback = function(score) {
          Huyen.sum += score;
          dailyScores.push(score);
        };
        switch(index) {
          case 0:
            echoAverageFrom(Dscores, 8);
            break;
          case 1:
            echoAverageFrom(Lscores, 8);
            break;
          case 2:
            echoAverageFrom(Tscores, 8);
            break;
          case 3:
            echoAverageFrom(ABscores, 8);
            break;
          default:
            break;
        }
      }

      if(name == Takuma.name) {
        Takuma.count++;
        var callback = function(score) {
          Takuma.sum += score;
          dailyScores.push(score);
        };

        switch(index) {
          case 0:
            echoAverageFrom(Dscores, 10);
            break;
          case 1:
            echoAverageFrom(Lscores, 10);
            break;
          case 2:
            echoAverageFrom(Tscores, 10);
            break;
          case 3:
            echoAverageFrom(ABscores, 10);
            break;
          default:
            break;
        }
      }

    });
    startRow++;
  }


  // A列に最新の要約文があるシートまで日付を出力
  var dateStartRow = 5;
  for(var key in sheetsObj) {
    evalationSheet.getRange(dateStartRow, 1).setValue(sheetsObj[key].getName());
    dateStartRow++;
  }

  memberObjs.map(function(member) {
    echoSumAndAverage(member);
  });

  function echoSumAndAverage(member) {
    evalationSheet.getRange(member.sumCell).setValue(member.sum);
    evalationSheet.getRange(member.aveCell).setValue(average(member));
  }

  function average(name) {
    return name.sum / name.count;
  }
}
