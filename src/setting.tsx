import React from 'react';
import { useSettingsButton, useCloudStorage, FieldPicker, useActiveViewId, useFields, useViewIds } from '@vikadata/widget-sdk';

export const Setting: React.FC = () => {
  const [isShowingSettings] = useSettingsButton()

  const viewIds = useViewIds()
  const activeViewId = useActiveViewId()
  const fields = useFields(activeViewId?activeViewId:viewIds[0])
  const [fieldId, setFieldId] = useCloudStorage<string>('selectedAttachmentFieldId', fields[0].id)

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

  return isShowingSettings ? (
    <div style={{ flexShrink: 0, width: '300px', borderLeft: 'solid 1px gainsboro', paddingLeft: '16px', paddingTop: '40px', paddingRight: '16px', backgroundColor: '#fff' }}>
      <h3>配置</h3>
      <div style={{ marginTop: '16px' }}>
        <label style={{ fontSize: '12px', color: '#999' }}>请选择 word 模板所在的附件列名</label>
        <FieldPicker viewId={activeViewId?activeViewId:viewIds[0]} fieldId={fieldId} onChange={option => checkAndUpdateSelectedAttachmentField(option.value)} />
      </div>
    </div>
  ) : null;
};
