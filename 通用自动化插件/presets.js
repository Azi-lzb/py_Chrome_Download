// ═══════════════════════════════════════════════
//  预置录制数据（Chrome DevTools Recorder 格式）
// ═══════════════════════════════════════════════

const PRESET_CR_RECORDINGS = [
  {
    name: 'FR筛选',
    data: {
      title: 'FR筛选',
      steps: [
        { type: 'navigate', url: 'http://11.210.4.110/fsdc2020/new/main/#/dataSharing/sharing005/appMaintain', assertedEvents: [{ type: 'navigation', url: 'http://11.210.4.110/fsdc2020/new/main/#/dataSharing/sharing005/appMaintain', title: '国家金融基础数据库大数据平台' }] },
        { type: 'click', target: 'main', selectors: [['div.fr-trigger-btn-over > div'], ['xpath//html/body/div[1]/div[1]/div[2]/div/div[15]/div[2]/div'], ['pierce/div.fr-trigger-btn-over > div']], offsetY: 6.031246185302734, offsetX: 14.0242919921875, frame: [1] },
        { type: 'click', target: 'main', selectors: [['body'], ['xpath//html/body'], ['pierce/body']], offsetY: 241, offsetX: 255, frame: [1] },
        { type: 'doubleClick', target: 'main', selectors: [['#content-container'], ['xpath///*[@id="content-container"]'], ['pierce/#content-container']], offsetY: 135.02777099609375, offsetX: 255, frame: [1] },
        { type: 'click', target: 'main', selectors: [['div.fr-trigger-btn-over'], ['xpath//html/body/div[1]/div[1]/div[2]/div/div[10]/div[2]'], ['pierce/div.fr-trigger-btn-over']], offsetY: 3.0173568725585938, offsetX: 0.020751953125, frame: [1] },
        { type: 'click', target: 'main', selectors: [['div.fr-trigger-btn-over'], ['xpath//html/body/div[1]/div[1]/div[2]/div/div[12]/div[2]'], ['pierce/div.fr-trigger-btn-over']], offsetY: 11.017356872558594, offsetX: 18.0242919921875, frame: [1] },
        { type: 'click', target: 'main', selectors: [['div.fr-combo-list-item > div'], ['xpath//html/body/div[3]/div[4]/div[2]/div'], ['pierce/div.fr-combo-list-item > div']], offsetY: 6.565971374511719, offsetX: 58.78125, frame: [1] },
        { type: 'click', target: 'main', selectors: [['div.fr-combo-list-item span'], ['xpath//html/body/div[3]/div[4]/div[2]/div/span'], ['pierce/div.fr-combo-list-item span'], ['text/广东省']], offsetY: 9.444442749023438, offsetX: 50.78125, frame: [1] },
        { type: 'click', target: 'main', selectors: [['div.parameter-container > div:nth-of-type(1)'], ['xpath//html/body/div[1]/div[1]/div[1]'], ['pierce/div.parameter-container > div:nth-of-type(1)']], offsetY: 2.013885498046875, offsetX: 913, frame: [1] },
        { type: 'click', target: 'main', selectors: [['div.fr-trigger-btn-over > div'], ['xpath//html/body/div[1]/div[1]/div[2]/div/div[13]/div[2]/div'], ['pierce/div.fr-trigger-btn-over > div']], offsetY: 0.031246185302734375, offsetX: 8.0103759765625, frame: [1] },
        { type: 'click', target: 'main', selectors: [['div:nth-of-type(5) > div.fr-combo-list-item span'], ['xpath//html/body/div[3]/div[5]/div[2]/div/span'], ['pierce/div:nth-of-type(5) > div.fr-combo-list-item span'], ['text/惠州市']], offsetY: 12.444442749023438, offsetX: 38.7777099609375, frame: [1] },
        { type: 'click', target: 'main', selectors: [['div.fr-trigger-btn-over > div'], ['xpath//html/body/div[1]/div[1]/div[2]/div/div[14]/div[2]/div'], ['pierce/div.fr-trigger-btn-over > div']], offsetY: 2.0312461853027344, offsetX: 10.017333984375, frame: [1] },
        { type: 'click', target: 'main', selectors: [['div:nth-of-type(6) > div.fr-checkbox-control span'], ['xpath//html/body/div[3]/div[6]/div[1]/div/span'], ['pierce/div:nth-of-type(6) > div.fr-checkbox-control span']], offsetY: 9.767356872558594, offsetX: 40.7916259765625, frame: [1] },
        { type: 'click', target: 'main', selectors: [['div.fr-trigger-btn-over'], ['xpath//html/body/div[1]/div[1]/div[2]/div/div[2]/div[2]'], ['pierce/div.fr-trigger-btn-over']], offsetY: 4.017356872558594, offsetX: 3.0242919921875, frame: [1] },
        { type: 'click', target: 'main', selectors: [['div.fr-trigger-btn-over'], ['xpath//html/body/div[1]/div[1]/div[2]/div/div[2]/div[2]'], ['pierce/div.fr-trigger-btn-over']], offsetY: 4.017356872558594, offsetX: 3.0242919921875, frame: [1] },
        { type: 'click', target: 'main', selectors: [['#content-container'], ['xpath///*[@id="content-container"]'], ['pierce/#content-container']], offsetY: 170.02777099609375, offsetX: 264, frame: [1] }
      ]
    }
  },
  {
    name: '任务导出',
    data: {
      title: '任务导出',
      steps: [
        { type: 'navigate', url: 'http://11.210.4.110/fsdc2020/new/main/#/bcs/taskSearch', assertedEvents: [{ type: 'navigation', url: 'http://11.210.4.110/fsdc2020/new/main/#/bcs/taskSearch', title: '国家金融基础数据库大数据平台' }] },
        { type: 'click', target: 'main', selectors: [['tr:nth-of-type(1) > td.el-table_2_column_29 button:nth-of-type(1) > span'], ['xpath///*[@id="avue-view"]/div/div[1]/div[2]/div[1]/div/div/div[3]/table/tbody/tr[1]/td[17]/div/button[1]/span'], ['pierce/tr:nth-of-type(1) > td.el-table_2_column_29 button:nth-of-type(1) > span']], offsetY: 12.027740478515625, offsetX: 38.798583984375 },
        { type: 'click', target: 'main', selectors: [['aria/查看报送说明'], ['div.avue-layout span > button'], ['xpath///*[@id="avue-view"]/div/div/div/div[1]/div[1]/div[2]/div/div/span/button'], ['pierce/div.avue-layout span > button']], offsetY: 25.784713745117188, offsetX: 47.555419921875 },
        { type: 'click', target: 'main', selectors: [['aria/生成校验结果文件', 'aria/[role="generic"]'], ['div.panel-header button:nth-of-type(1) > span'], ['xpath///*[@id="avue-view"]/div/div/div/div[1]/div[1]/div[2]/button[1]/span'], ['pierce/div.panel-header button:nth-of-type(1) > span'], ['text/生成校验结果文件']], offsetY: 13.666656494140625, offsetX: 79.7291259765625 },
        { type: 'click', target: 'main', selectors: [['aria/导出校验结果文件', 'aria/[role="generic"]'], ['div.panel-header button:nth-of-type(2) > span'], ['xpath///*[@id="avue-view"]/div/div/div/div[1]/div[1]/div[2]/button[2]/span'], ['pierce/div.panel-header button:nth-of-type(2) > span']], offsetY: 11.666656494140625, offsetX: 45.586669921875 },
        { type: 'click', target: 'main', selectors: [['aria/导出校验结果文件'], ['div.panel-header button:nth-of-type(2)'], ['xpath///*[@id="avue-view"]/div/div/div/div[1]/div[1]/div[2]/button[2]'], ['pierce/div.panel-header button:nth-of-type(2)']], offsetY: 22.548599243164062, offsetX: 38.6978759765625, assertedEvents: [{ type: 'navigation' }] },
        { type: 'click', target: 'main', selectors: [['aria/查看报送说明 [role="tab"]', 'aria/[role="generic"]'], ['xpath///*[@id="tab-/bcs/viewReportExplain_ab63953859f04fdaae49d84a097599c2_all?includeVerifyResult=0&missionInstanceId=ab63953859f04fdaae49d84a097599c2&missionTemplateName=%E9%87%91%E8%9E%8D%E5%9F%BA%E7%A1%80%E6%95%B0%E6%8D%AE%E9%87%87%E9%9B%86%EF%BC%88%E4%BB%A3%E5%8F%91%E5%B7%A5%E8%B5%84%EF%BC%89&isCustom=0&isAudit=false&pkgId=&fileName=&missionInstanceStatus=08&isSingle=0&parentFrom=search&viewMode=view&batchNum=&dataDate=2026-04-30&fileType=\"]/span'], ['text/查看报送说明']], offsetY: 7.166664123535156, offsetX: 10.42706298828125 },
        { type: 'click', target: 'main', selectors: [['aria/导出核查表', 'aria/[role="generic"]'], ['button:nth-of-type(3) > span'], ['xpath///*[@id="avue-view"]/div/div/div/div[1]/div[2]/div/div/div[2]/div[2]/div/form/div[2]/div/div/div/div/button[3]/span'], ['pierce/button:nth-of-type(3) > span'], ['text/导出核查表']], offsetY: 0.8055419921875, offsetX: 26.46875 },
        { type: 'click', target: 'main', selectors: [['aria/确定', 'aria/[role="generic"]'], ['div.el-message-box__wrapper button.el-button--primary > span'], ['xpath//html/body/div[7]/div/div[3]/button[2]/span'], ['pierce/div.el-message-box__wrapper button.el-button--primary > span']], offsetY: 6.1666259765625, offsetX: 13.8818359375, assertedEvents: [{ type: 'navigation' }] },
        { type: 'click', target: 'main', selectors: [['aria/查看任务 [role="tab"]', 'aria/[role="generic"]'], ['text/查看任务']], offsetY: 12.166664123535156, offsetX: 6.6180419921875 },
        { type: 'click', target: 'main', selectors: [['aria/查看报送说明'], ['div.avue-layout span > button'], ['xpath///*[@id="avue-view"]/div/div/div/div[1]/div[1]/div[2]/div/div/span/button'], ['pierce/div.avue-layout span > button']], offsetY: 5.7847137451171875, offsetX: 71.555419921875 },
        { type: 'click', target: 'main', selectors: [['aria/生成校验结果文件'], ['div.panel-header button:nth-of-type(1)'], ['xpath///*[@id="avue-view"]/div/div/div/div[1]/div[1]/div[2]/button[1]'], ['pierce/div.panel-header button:nth-of-type(1)']], offsetY: 4.5485992431640625, offsetX: 43.8402099609375 },
        { type: 'click', target: 'main', selectors: [['aria/导出校验结果文件', 'aria/[role="generic"]'], ['div.panel-header button:nth-of-type(2) > span'], ['xpath///*[@id="avue-view"]/div/div/div/div[1]/div[1]/div[2]/button[2]/span'], ['pierce/div.panel-header button:nth-of-type(2) > span'], ['text/导出校验结果文件']], offsetY: 13.666656494140625, offsetX: 59.586669921875, assertedEvents: [{ type: 'navigation' }] },
        { type: 'click', target: 'main', selectors: [['aria/查看报送说明 [role="tab"]', 'aria/[role="generic"]'], ['text/查看报送说明']], offsetY: 7.166664123535156, offsetX: 8.42706298828125 },
        { type: 'click', target: 'main', selectors: [['aria/导出核查表'], ['#app button:nth-of-type(3)'], ['xpath///*[@id="avue-view"]/div/div/div/div[1]/div[2]/div/div/div[2]/div[2]/div/form/div[2]/div/div/div/div/button[3]'], ['pierce/#app button:nth-of-type(3)']], offsetY: 5.6875, offsetX: 70.486083984375 },
        { type: 'click', target: 'main', selectors: [['aria/确定', 'aria/[role="generic"]'], ['div.el-message-box__wrapper button.el-button--primary > span'], ['xpath//html/body/div[7]/div/div[3]/button[2]/span'], ['pierce/div.el-message-box__wrapper button.el-button--primary > span']], offsetY: 1.1666259765625, offsetX: 3.8818359375, assertedEvents: [{ type: 'navigation' }] },
        { type: 'click', target: 'main', selectors: [['aria/查看任务 [role="tab"]', 'aria/[role="generic"]'], ['text/查看任务']], offsetY: 7.166664123535156, offsetX: 9.6180419921875 }
      ]
    }
  }
];
