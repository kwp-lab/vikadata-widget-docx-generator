import React from 'react';
import { useSettingsButton, useCloudStorage, FieldPicker, useActiveViewId, useFields, useViewIds, FieldType } from '@vikadata/widget-sdk';
import { ISwitchProps, Switch } from '@vikadata/components';
import { InformationSmallOutlined } from '@vikadata/icons';

export const Setting: React.FC = () => {
  const [isShowingSettings] = useSettingsButton()

  const viewIds = useViewIds()
  const activeViewId = useActiveViewId()
  const fields = useFields(activeViewId?activeViewId:viewIds[0])
  const [fieldId, setFieldId] = useCloudStorage<string>('selectedAttachmentFieldId', fields[0].id)
  const [keepFormat, setKeepFormat] = useCloudStorage<boolean>('keepFormat', true)

  const checkAndUpdateSelectedAttachmentField = function(selectedFieldId:string){
    fields.forEach(field => {
      if(field.id == selectedFieldId){
        if(field.type == 'Attachment'){
          setFieldId(selectedFieldId)
        }else{
          alert("请选择一个附件类型的字段！")
        }
      }
    })
  }

  const keepFormatSwitchProps:ISwitchProps = {
    onChange: (value) => setKeepFormat(value)
  }

  if(keepFormat) keepFormatSwitchProps.checked = true;

  return isShowingSettings ? (
    <div style={{ flexShrink: 0, width: '300px', borderLeft: 'solid 1px var(--lineColor)', paddingLeft: '16px', paddingTop: '40px', paddingRight: '16px', background: 'var(--defaultBg)' }}>
      <h3 style={{color: 'var(--firstLevelText)'}}>
        配置 
        <a style={{verticalAlign: "middle", color: 'var(--thirdLevelText)', marginLeft: "4px"}} title="查看教程" target="_blank" href="https://bbs.vika.cn/article/111" >
          <InformationSmallOutlined size={16}  />
        </a>
      </h3>
      <div style={{ marginTop: '16px' }}>
        <label style={{ fontSize: '12px', color: 'var(--thirdLevelText)' }}>请选择 word 模板所在的附件列名</label>
        <div style={{background: 'var(--fill0)'}}>
          <FieldPicker 
            viewId={activeViewId?activeViewId:viewIds[0]} 
            fieldId={fieldId} 
            onChange={option => checkAndUpdateSelectedAttachmentField(option.value)}
            allowedTypes={[FieldType.Attachment]}
          />
        </div>
      </div>
      <div style={{ marginTop: '16px' }}>
        <label style={{ fontSize: '12px', color: 'var(--thirdLevelText)' }}>保持与单元格的显示格式一致</label>
        <div>
          <Switch {...keepFormatSwitchProps} />
          <p style={{
                color: 'var(--thirdLevelText)',
                fontSize: '12px',
                padding: '4px',
                border: '1px dashed #ccc',
                borderRadius: '6px',
                marginTop: '0.5em'
          }}>
            输出到word文档里的内容，会保持跟表格里看到的一致，以货币类型举例：<br/>
            开启前：1200<br/>
            开启后："$1,200.00"<br/>
            本配置项仅对货币，数字，百分比，日期列类型生效
          </p>
        </div>
      </div>
    </div>
  ) : null;
};
