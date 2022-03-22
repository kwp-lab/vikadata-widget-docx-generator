import React from 'react';
import { useDatasheet } from '@vikadata/widget-sdk';
import { DocxGenerator } from './docx_generator';
import { Setting } from './setting';

export const WidgetDeveloperTemplate: React.FC = () => {

  const datasheet = useDatasheet()
  
  // 校验用户是否有新增记录的权限，从而判断用户对表格是否只读权限
  const permission = datasheet?.checkPermissionsForAddRecord()

  return (
    <div style={{ display: 'flex', height: '100%', backgroundColor: '#fff', borderTop: '1px solid gainsboro' }}>
      <div style={{ flexGrow: 1, overflow: 'auto', padding: '0 8px'}}>
        <DocxGenerator />
      </div>
      {permission?.acceptable && <Setting/>}
      
    </div>
  );
};
