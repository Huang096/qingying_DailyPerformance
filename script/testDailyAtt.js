// Copyright 2024 huangzheheng
// 
// Licensed under the Apache License, Version 2.0 (the "License");
// you may not use this file except in compliance with the License.
// You may obtain a copy of the License at
// 
//     https://www.apache.org/licenses/LICENSE-2.0
// 
// Unless required by applicable law or agreed to in writing, software
// distributed under the License is distributed on an "AS IS" BASIS,
// WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
// See the License for the specific language governing permissions and
// limitations under the License.


// 这里的代码用来测试 “判断当天出勤” 的情况。
// 测试条件 checkInTime.getHours() < 9 && checkOutTime.getHours() >= 14
// 仅仅是用来判断特定的 “一天”，出勤情况为正常记录为1，不正常记录为0
// 但是这里的测试代码会把之前的记录全部删除掉，只保留唯一一条测试记录


function testOnFormSubmit() {
    // 获取源表格和目标表格
    var sourceSheetId = '1a97MwNd5rrZuBc9m6oGJIxeuhcs5KQyi0mx5Ylu-hBM';
    var sourceSheetName = '签到情况';

    var sourceSpreadsheet = SpreadsheetApp.openById(sourceSheetId);
    var sourceSheet = sourceSpreadsheet.getSheetByName(sourceSheetName);
    var targetSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('签到总结');

    // 清空源表格和目标表格的现有数据
    sourceSheet.clear();
    targetSheet.clear();

    // 添加表头
    sourceSheet.appendRow(['时间戳', '姓名', '状态']);
    targetSheet.appendRow(['姓名', '签到时间', '签退时间', '工作时间', '出勤']);

    // 添加测试数据
    sourceSheet.appendRow([new Date('2024-06-13T09:45:00'), '1', '签到']); 
    sourceSheet.appendRow([new Date('2024-06-13T16:30:00'), '1', '签退']); 

    // 模拟表单提交事件
    var e = { values: [new Date(), '测试用户', '签到'] };
    onFormSubmit(e);
  }
  
  function onFormSubmit(e) {
    Logger.log("Form submitted");
  
    // 第一个表格的实际 ID 和工作表名称
    var sourceSheetId = '1a97MwNd5rrZuBc9m6oGJIxeuhcs5KQyi0mx5Ylu-hBM';
    var sourceSheetName = '签到情况';
  
    var sourceSpreadsheet = SpreadsheetApp.openById(sourceSheetId);
    var sourceSheet = sourceSpreadsheet.getSheetByName(sourceSheetName);
  
    if (!sourceSheet) {
      Logger.log("Source Sheet not found!");
      return;
    }
  
    var targetSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('签到总结');
  
    if (!targetSheet) {
      Logger.log("Target Sheet not found!");
      return;
    }
  
    var sourceData = sourceSheet.getDataRange().getValues();
    var targetData = targetSheet.getDataRange().getValues();
  
    // 创建一个记录签到状态的对象
    var checkInRecords = {};
  
    // 遍历第一个表格的数据，查找成对的签到和签退记录
    for (var i = 1; i < sourceData.length; i++) {
      var timeStamp = new Date(sourceData[i][0]);
      var name = sourceData[i][1];
      var status = sourceData[i][2];
  
      if (status == "签到") {
        checkInRecords[name] = checkInRecords[name] || {};
        checkInRecords[name].checkInTime = timeStamp;
      } else if (status == "签退") {
        checkInRecords[name] = checkInRecords[name] || {};
        checkInRecords[name].checkOutTime = timeStamp;
      }
    }
  
    Logger.log("Check-in Records: " + JSON.stringify(checkInRecords));
  
    // 更新第二个表格
    for (var name in checkInRecords) {
      if (checkInRecords[name].checkInTime && checkInRecords[name].checkOutTime) {
        var checkInTime = checkInRecords[name].checkInTime;
        var checkOutTime = checkInRecords[name].checkOutTime;
        var workTime = (checkOutTime - checkInTime) / (1000 * 60 * 60); // 计算小时数
        // 记录当天出勤情况——黄哲亨
        var attendance = 0;
        if (checkInTime.getHours() < 9 && checkOutTime.getHours() >= 14) {
          attendance = 1;
        }
        
        // 在第二个表格中找到对应的行
        var rowIndex = targetData.findIndex(function(row) {
          return row[0] == name && row[2] == "";
        });
        
        if (rowIndex == -1) {
          // 如果没有找到对应的行，则添加一行
          // 多加一项出勤——黄哲亨
          targetSheet.appendRow([name, checkInTime, checkOutTime, workTime.toFixed(2) + " 小时", attendance]);
          Logger.log("Appended new row with attendance: " + attendance);
        } else {
          // 更新现有的行
          targetSheet.getRange(rowIndex + 1, 2).setValue(checkInTime);
          targetSheet.getRange(rowIndex + 1, 3).setValue(checkOutTime);
          targetSheet.getRange(rowIndex + 1, 4).setValue(workTime.toFixed(2) + " 小时");
          // 写入出勤——黄哲亨
          targetSheet.getRange(rowIndex + 1, 5).setValue(attendance);
          Logger.log("Updated row " + (rowIndex + 1) + " with attendance: " + attendance);
        }
      }
    }

    // 删除已经处理的原始数据行
    for (var i = sourceData.length - 1; i > 0; i--) {
        var name = sourceData[i][1];
        var status = sourceData[i][2];
        if (checkInRecords[name] && checkInRecords[name].checkInTime && checkInRecords[name].checkOutTime) {
            sourceSheet.deleteRow(i + 1);
        }
    }
}
