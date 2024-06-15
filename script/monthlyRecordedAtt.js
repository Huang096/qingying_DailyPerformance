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

// 这里的代码是基于 我的打卡记录是否符合出勤要求，但是这仅仅局限于“存在的打卡行为”。如果我今天
// 没有打卡，这个是暂时没办法判断的。关于这个怎么判断，我想到有两个方法。
// 1. (功能上可能合理的)算出当月应打卡天数。比如说可以写一个函数获取当月所有工作日的天数，然后
// 追踪打卡人本月打卡天数，然后比较。但是这个东西要等到月底才能比较。
// 2. (我觉得合理的)记录每个人的 “本月打卡天数”，然后每天都能够直接看出来大家打卡的天数，月底
// 结算的时候也挺一目了然的其实。

// 我觉得老板的使用场景应 该是：今天心情不错，让我打开本月出勤表来看看大家到今天为止的出勤情况
// 吧。然后我们的使用场景是：好累，今天好想摸鱼，但是看到了邮件提醒，手贱点进来看到大家都在努力
// 工作，如果我今天不上班打卡，老板就会看到我工作天数比别人少，就会看到我划水，就不给我return
// offer，就会扣钱。所以我一定要认真上班！
// 所以我觉得实时记录天数很不错




function onFormSubmit(e) {
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
  
    // 判断当月是否已经有表来记录，如果没有那就新建——黄哲亨
    //var today = new Date;
    //用来测试不同月份是否能成功建表, 测试通过
    var today = new Date('2024-07-15');
    var month = today.getMonth() + 1;
    var year = today.getFullYear();
    var monthlySheetName = year + '-' + month + ' 签到情况';
  
    var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    if (!spreadsheet.getSheetByName(monthlySheetName)) {
      createMonthlySheet(monthlySheetName);
    }
  
  
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
        } else {
          // 更新现有的行
          targetSheet.getRange(rowIndex + 1, 2).setValue(checkInTime);
          targetSheet.getRange(rowIndex + 1, 3).setValue(checkOutTime);
          targetSheet.getRange(rowIndex + 1, 4).setValue(workTime.toFixed(2) + " 小时");
          // 写入出勤——黄哲亨
          targetSheet.getRange(rowIndex + 1, 5).setValue(attendance);
        }
        updateMonthlySheet(monthlySheetName, name, attendance);
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
  
  // 创建每个月的签到情况表——黄哲亨
  function createMonthlySheet(sheetName) {
    var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  
    // 检查是否已存在
    if (spreadsheet.getSheetByName(sheetName)) {
      Logger.log("Sheet " + sheetName + " already exists");
      return;
    }
  
    // 创建新的月份表格
    var newSheet = spreadsheet.insertSheet(sheetName);
    newSheet.appendRow(['姓名', '本月出勤']);
    Logger.log("Created new sheet: " + sheetName);
  }
  
  //当打卡人签退的时候，当日出勤情况不正常(为0),更新当月出勤情况。默认情况是全勤
  function updateMonthlySheet(monthlySheetName, name, attendance) {
    var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    var monthlySheet = spreadsheet.getSheetByName(monthlySheetName);
    var monthlyData = monthlySheet.getDataRange().getValues();
    
    var monthlyRowIndex = monthlyData.findIndex(function(row) {
      return row[0] == name;
    });
    
    if (monthlyRowIndex == -1) {
      // 如果没有找到对应的行，则添加一行
      monthlySheet.appendRow([name, '全勤']);
    } else {
      // 如果本次出勤状态为0，且记录存在并且目前是全勤，则更新为有缺勤记录
      var currentAttendance = monthlySheet.getRange(monthlyRowIndex + 1, 2).getValue();
      if (attendance == 0 && currentAttendance == '全勤') {
        monthlySheet.getRange(monthlyRowIndex + 1, 2).setValue('有缺勤记录');
      }
    }
  }
  
  
  function onOpen() {
    var ui = SpreadsheetApp.getUi();
    ui.createMenu('考勤系统')
      .addItem('处理考勤记录', 'onFormSubmit')
      .addToUi();
  }
  
  function testOnFormSubmit() {
    var sourceSheetId = '1a97MwNd5rrZuBc9m6oGJIxeuhcs5KQyi0mx5Ylu-hBM';
    var sourceSheetName = '签到情况';
    var sourceSpreadsheet = SpreadsheetApp.openById(sourceSheetId);
    var sourceSheet = sourceSpreadsheet.getSheetByName(sourceSheetName);
  
    // 清空源表格的现有数据
    sourceSheet.clear();
    sourceSheet.appendRow(['时间戳', '姓名', '状态']);
  
    // 添加测试数据
    var testData = [
      ['2024-07-15 08:30:00', '张三', '签到'],
      ['2024-07-15 17:30:00', '张三', '签退'],
      ['2024-07-15 09:30:00', '李四', '签到'],
      ['2024-07-15 17:30:00', '李四', '签退'],
      ['2024-07-15 08:30:00', '王五', '签到'],
      ['2024-07-15 13:30:00', '王五', '签退'],
      ['2024-07-15 09:30:00', '赵六', '签到'],
      ['2024-07-15 18:30:00', '赵六', '签退'],
    ];
  
    testData.forEach(function(row) {
      sourceSheet.appendRow(row);
    });
  
    // 模拟表单提交事件
    var e = { values: [new Date(), '测试用户', '签到'] };
    onFormSubmit(e);
  }
  