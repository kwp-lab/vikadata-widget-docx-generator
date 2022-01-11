import React from 'react';
import { DocxGenerator } from './docx_generator';
import { Setting } from './setting';

export const WidgetDeveloperTemplate: React.FC = () => {
  return (
    <div style={{ display: 'flex', height: '100%', backgroundColor: '#fff', borderTop: '1px solid gainsboro' }}>
      <div style={{ flexGrow: 1, overflow: 'auto', padding: '0 8px'}}>
        <DocxGenerator />
      </div>
      <Setting/>
    </div>
  );
};
