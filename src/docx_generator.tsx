import { useRecords, useFields, useActiveViewId, useSelection, useCloudStorage, useSettingsButton, useViewport, useField, Field, Record, IAttachmentValue, usePrimaryField } from '@vikadata/widget-sdk';
import { AttachmentPasteOutlined, InfoFilled } from '@vikadata/icons';
import { Button } from '@vikadata/components';
import React from 'react';
import Docxtemplater from 'docxtemplater';
import PizZip from 'pizzip';
import PizZipUtils from 'pizzip/utils/index.js';
import { saveAs } from 'file-saver';

const userToken = ""

/**
 * 通过URL读取文件内容
 */
function loadFile(url: String, callback) {
  PizZipUtils.getBinaryContent(url, callback);
}

function replaceErrors(key: String, value: any) {
  if (value instanceof Error) {
    return Object.getOwnPropertyNames(value).reduce(function (
      error,
      key
    ) {
      error[key] = value[key];
      return error;
    },
      {});
  }
  return value;
}

/**
 * 将 Blob 文件对象作为附件上传到指定的表格
 * @param activeDatasheetId 当前表格的ID
 * @param fileBlob 待上传的文件对象
 */
async function uploadAttachment(activeDatasheetId: String, fileBlob: Blob) {
  console.log("uploadAttachment", activeDatasheetId)

  const url = `https://api.vika.cn/fusion/v1/datasheets/${activeDatasheetId}/attachments`
  let formData = new FormData()
  let file = new File([fileBlob], "tag-example.docx", { "type": "'application/vnd.openxmlformats-officedocument.wordprocessingml.document'" })
  formData.append('file', file);

  return await fetch(url, {
    method: "POST",
    body: formData,
    headers: {
      'Authorization': `Bearer ${userToken}`
    }
  }).then(res => res.json())
}

/**
 * 文档生成的异常处理
 * @param error 
 */
function throwError(error: any) {
  console.log(JSON.stringify({ error: error }, replaceErrors));

  if (error.properties && error.properties.errors instanceof Array) {
    const errorMessages = error.properties.errors
      .map(function (error) {
        return error.properties.explanation;
      })
      .join('\n');
    console.log('errorMessages', errorMessages);
  }
  throw error;
}

/**
 * 遍历已选择的多条 record ，从中获取数据并生成 word 文档
 */
function generateDocuments(selectedRecords: Record[], fields: Field[], selectedAttachmentField: Field, primaryField: Field) {

  for (let index = 0; index < selectedRecords.length; index++) {
    const record = selectedRecords[index]
    const row = {}
    const filename = record.getCellValueString(primaryField.id) || "未命名"

    fields.forEach(field => {
      row[field.name] = record.getCellValueString(field.id) || "(空值)"
    })

    const attachements = record.getCellValue(selectedAttachmentField.id)
    const attachmentName = attachements[0].name

    console.log({ row, attachements })

    const prefix = attachmentName.substr(0, attachmentName.lastIndexOf("."))

    attachements && generateDocument(row, attachements[0], prefix + "-" + filename)
  }

}

/**
 * 调用第三方库，生成word文档并调起浏览器附件下载事件
 */
function generateDocument(row: any, selectedAttachment: IAttachmentValue, filename: string) {

  loadFile(selectedAttachment.url, function (error, content) {
    if (error) {
      throw error
    }

    const zip = new PizZip(content)

    const doc = new Docxtemplater(zip, {
      paragraphLoop: true,
      linebreaks: true
    })

    doc.setData(row)

    try {
      doc.render();
    } catch (error: any) {
      throwError(error)
    }

    // const out = doc.getZip().generate({
    //   type: 'blob',
    //   mimeType:
    //     'application/vnd.openxmlformats-officedocument.wordprocessingml.document'
    // });

    const out = doc.getZip().generate({
      type: 'blob',
      mimeType:
        'application/vnd.openxmlformats-officedocument.wordprocessingml.document'
    });

    saveAs(out, filename + ".docx")

    // await uploadAttachment(activeDatasheetId, out).then(res => {
    //   console.log(res)
    // })


  });
}

/**
 * 小程序展开状态下，显示 Readme 信息
 */
function showReadmeInfo() {
  const wrapperStyle: React.CSSProperties = {
    width: '100%',
    padding: '10px'
  }

  const imgStyle: React.CSSProperties = {
    width: '100%',
    border: '1px solid #9484f1',
    borderRadius: '4px'
  }


  return (
    <div style={wrapperStyle}>
      <h1>Word文档生成器</h1>
      <h3><b>前言</b></h3>
      <p>你有遇到过下面这些状况吗？<br />日复一日地填写多份相同格式的 Word 文档...<br />经常因为 Word 文档搬运而加班到深夜...</p>
      <p>试试本小组件吧！让你从重复性的 Word 搬运工中解放出来！^0^</p>

      <h3><b>简介</b></h3>
      <p>本小组件可以将每一行数据填充到 Word 模板里面，从而形成一份新的 Word 文档。同时选中多行记录，即可实现批量导出 Word 文档。</p>
      <p>例如一份《录取通知书》。在日常工作中，公司HR一天可能会发送多份《录取通知书》，里面的格式都是一样的，只是“岗位”，“部门”，“候选人姓名”，“通知日期”等等这些信息要素会有所不同，但HR却需要手工重复性地复制粘贴、复制粘贴...</p>
      <p>使用本小组件后，只需要提前制作一次 Word 模板，往后的工作就只需要点一点手指头，小组件来帮你填充关键信息要素，并生成新的《录取通知书》！</p>

      <h3><b>使用步骤</b></h3>
      <p>1. 提前准备好 Word 模板，在 Word 模板里面目标位置填写好维格表里的对应列名，写法跟智能公式里引用单元格值一样，在列名左右两边加上花括号，例如“<code>{'\u007B'}候选人姓名{'\u007D'}</code>”</p>
      <p>2. 将修改好的 Word 模板以附件形式上传到当前维格表的附件列里，如下图示例</p>
      <p><img src="https://s1.vika.cn/space/2021/12/02/22202756884f485dbfce5e257000644c" alt="示意图" style={imgStyle} /></p>
      <p>3. 在本界面右侧的配置区域选择 Word 模板所在的附件列名</p>
      <p>4. 点击右上角按钮，退出小组件的“展开模式”</p>
      <p>5. 在维格视图中选择若干行，然后点击小组件的“导出 Word 文档”</p>
      <p style={{ textAlign: 'center' }} >------ 至此，可以开启高效办公之旅了 :） ------</p>
    </div>
  )
}

export const DocxGenerator: React.FC = () => {
  const { isFullscreen, toggleFullscreen } = useViewport()
  const [isShowingSettings, toggleSettings] = useSettingsButton()
  const activeViewId = useActiveViewId()
  const selection = useSelection()
  const selectedRecords = useRecords(activeViewId, { ids: selection?.recordIds })
  const fields = useFields(activeViewId)
  const primaryField = usePrimaryField() || fields[0]

  const [fieldId] = useCloudStorage<string>('selectedAttachmentFieldId')
  const selectedAttachmentField = useField(fieldId)

  const openSettingArea = function () {
    !isFullscreen && toggleFullscreen()
    !isShowingSettings && toggleSettings()
  }

  if (isFullscreen) {
    return showReadmeInfo()
  }

  const style1 = {
    display: 'flex',
    alignContent: 'center',
    justifyContent: 'center',
    alignItems: 'center',
    height: '100%'
  }

  return (
    <div style={style1}>


      {selectedAttachmentField &&
        <div>
          <div style={{
            display: 'flex',
            alignContent: 'center',
            justifyContent: 'center',
            alignItems: 'center',
            width: '100%'
          }}>
            <AttachmentPasteOutlined size="50%" color="#7B67EE" />
          </div>

          <div style={{ marginTop: '16px', textAlign: 'center' }}>
            {(selectedRecords.length > 0) && <div>已选中 <span style={{ color: '#E33E38', fontWeight: 'bold' }}>{selectedRecords.length}</span> 条记录</div>}
          </div>

          <div style={{ marginTop: '16px', textAlign: 'center' }}>
            <Button onClick={generateDocuments.bind(this, selectedRecords, fields, selectedAttachmentField, primaryField)} variant="fill" color="primary" >导出 Word 文档</Button>
          </div>
        </div>
      }

      {!selectedAttachmentField &&
        <div>
          <div style={{
            display: 'flex',
            alignContent: 'center',
            justifyContent: 'center',
            alignItems: 'center',
            width: '100%'
          }}>
            <InfoFilled size="30%" />
          </div>
          <div>请设置一个存储word模板的附件字段</div>

          <div style={{ marginTop: '16px', textAlign: 'center' }}>
            <Button onClick={() => openSettingArea()} variant="fill" color="primary" >前往设置</Button>
          </div>
        </div>
      }
    </div>
  );
};
