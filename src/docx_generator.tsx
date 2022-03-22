import { useRecords, useFields, useActiveViewId, useSelection, useCloudStorage, useSettingsButton, useViewport, useField, Field, Record, IAttachmentValue, usePrimaryField, FieldType, useDatasheet } from '@vikadata/widget-sdk';
import { Button } from '@vikadata/components';
import { InformationSmallOutlined } from '@vikadata/icons';
import React, { useEffect, useState } from 'react';
import Docxtemplater from 'docxtemplater';
import {DXT} from 'docxtemplater';
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
  console.log("模板解析错误", JSON.stringify({ error: error }, replaceErrors));

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
      row[field.name] = record.getCellValue(field.id) || ""
      if(field.type == FieldType.MagicLink){
        // TODO
      } else if(field.type == FieldType.MultiSelect){
        row[field.name] = record.getCellValue(field.id) || []
        row[field.name] = row[field.name].map(item => {
          return item.name
        })
      }
    })

    

    const attachements = record.getCellValue(selectedAttachmentField.id)
    if (!attachements) {
      alert(`在指定的附件字段中找不到word模板，请上传。record:[${filename}]`)
      break
    }
    const attachmentName = attachements[0].name

    const prefix = attachmentName.substr(0, attachmentName.lastIndexOf("."))
    const suffix = attachmentName.substr(attachmentName.lastIndexOf(".")).toLowerCase()

    if(suffix !== ".docx"){
      alert(`只支持.docx格式的word模板（当前模板：${attachmentName}）`)
      break
    }

    console.log({ row, attachements, prefix, suffix })

    attachements && generateDocument(row, attachements[0], prefix + "-" + filename)
  }

}

/**
 * Docxtemplater 自定义标签解析器
 * @param tag 标签名称，eg: {产品名称} 
 * @returns 
 */
 const customParser = (tag) => {
  const isTernaryReg = new RegExp(/(.*)\?(.*)\:(.*)/)

  // 这是一个三元表达式
  const TernaryResult = isTernaryReg.exec(tag)
  var data1 = "";
  var data2 = "";
  if(TernaryResult !== null){
    tag = TernaryResult[1]
    data1 = TernaryResult[2]
    data2 = TernaryResult[3]
  }

  return {
    get(scope, context: DXT.ParserContext) {
      console.log({ tag, scope, context, TernaryResult })

      if (["$index", "$序号"].includes(tag)) {
        const indexes = context.scopePathItem
        return indexes[indexes.length - 1] + 1
        
      } else if(tag === "$isLast"){
        const totalLength = context.scopePathLength[context.scopePathLength.length - 1]
        const index = context.scopePathItem[context.scopePathItem.length - 1]
        return index === totalLength - 1

      } else if(tag == "$isFirst"){
        const index = context.scopePathItem[context.scopePathItem.length - 1]
        return index === 0

      } else if(tag.match(/(.*)\|find\((.*)\)/) !== null) {
        let [, fieldName, valueToFind] = tag.match(/(.*)\|find\((.*)\)/)
        fieldName = fieldName.trim()
        valueToFind = valueToFind.trim()
        console.log("detect find()", [fieldName, valueToFind, scope])
        if(fieldName && valueToFind && scope[fieldName] && Array.isArray(scope[fieldName])){
          const result =scope[fieldName].find(arrayItem => {
            if(typeof arrayItem == "string"){
              return (arrayItem==valueToFind) ? true : false
            }
            return false
          })
          return result ? true : false
        }
      } else if( tag.indexOf("==")>0 ){

          let [leftVal, rightVal] = tag.split("==")
          leftVal = leftVal.trim()
          rightVal = rightVal.trim().replace(/(“|”|’|‘|"|')/g, '')
          console.log("比较", {tag, leftVal, rightVal, scopeLeftVal: scope[leftVal], TernaryResult})
          if(TernaryResult !== null){
            return (scope[leftVal] === rightVal) ? TernaryResult[2] : TernaryResult[3]
          }else{
            return (scope[leftVal] == rightVal) ? scope[leftVal] : ""
          }
      }
      return scope[tag]
    }
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

    try {

      const doc = new Docxtemplater(zip, {
        paragraphLoop: true,
        linebreaks: true,
        parser: customParser,
      })

      doc.setData({...row, userGreeting: (scope) => {
        return "How is it going, " + scope.user + " ? ";
    }})

      try {
        doc.render();
      } catch (error: any) {
        throwError(error)
      }

      const out = doc.getZip().generate({
        type: 'blob',
        mimeType:
          'application/vnd.openxmlformats-officedocument.wordprocessingml.document'
      });

      saveAs(out, filename + ".docx")
    } catch (error) {
      console.log("错误信息", error)
      alert(`文件 ${selectedAttachment.name} 的模板语法不正确，请检查`)
    }


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
    width: "100%",
    padding: "10px 20px",
    display: "flex",
    alignItems: "center",
    height: "100%",
    justifyContent: "center",
    flexDirection: "column"
  }

  return (
    <div style={wrapperStyle}>
      <div style={{fontSize: "1.5em", color: "#C8C8C8", marginBottom: "0.5em"}}>不支持在小程序展开状态下导出word文档</div>
      <div>
        <a style={{fontSize: "16px"}} href="https://bbs.vika.cn/article/111" target="_blank" >
          <span style={{verticalAlign: "middle", lineHeight: "16px"}}>
            <InformationSmallOutlined size={16} color="#7b67ee" />
          </span> 查看教程
        </a>
      </div>
    </div>
  )
}

export const DocxGenerator: React.FC = () => {
  const { isFullscreen, toggleFullscreen } = useViewport()
  const [isShowingSettings, toggleSettings] = useSettingsButton()

  const activeViewId = useActiveViewId()
  const selection = useSelection()
  const selectionRecords = useRecords(activeViewId, { ids: selection?.recordIds })
  const fields = useFields(activeViewId)
  const primaryField = usePrimaryField() || fields[0]

  const datasheet = useDatasheet()
  
  // 校验用户是否有新增记录的权限，从而判断用户对表格是否只读权限
  const permission = datasheet?.checkPermissionsForAddRecord()


  // 读取配置
  const [fieldId] = useCloudStorage<string>('selectedAttachmentFieldId')
  const selectedAttachmentField = useField(fieldId)

  const [selectedRecords, setSelectedRecords] = useState<Record[]>([])

  const recordIds = selectionRecords.map((record: Record)=>{
    return record.recordId
  }).join(",")

  console.log("selectionRecords", selectionRecords)

  useEffect(() => {
    console.log({selectionRecords})
    if(Array.isArray(selectionRecords) && selectionRecords.length>0){
      setSelectedRecords(selectionRecords)
    }
  }, [selectionRecords.length, recordIds])

  const openSettingArea = function () {
    if(permission?.acceptable){
      !isFullscreen && toggleFullscreen()
      !isShowingSettings && toggleSettings()
    }else{
      alert("抱歉，只读权限无法进行此操作")
    }
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

  const helpLink = (
    <a style={{verticalAlign: "middle", position: "absolute", bottom: "8px", right: "8px", fontSize: "12px", color: "#8C8C8C"}} title="查看教程" target="_blank" href="https://bbs.vika.cn/article/111"  >
      <span style={{verticalAlign: "middle", lineHeight: "16px"}}><InformationSmallOutlined size={12} color="#8C8C8C" /></span>
      <span> 教程</span>
    </a>
  )

  return (
    <div style={style1}>
      {helpLink} 

      {selectedAttachmentField &&
        <div>
          <div style={{
            display: 'flex',
            alignContent: 'center',
            justifyContent: 'center',
            alignItems: 'center',
            width: '100%'
          }}>
            <img src='https://s1.vika.cn/space/2021/12/29/ce15dd51bb79495ab0f03ddf40d6fe92' style={{ width: '40%' }} />
          </div>

          <div style={{ textAlign: 'center' }}>
            {(selectedRecords.length > 0) && <div>已选中 <span style={{ color: '#fb4a43', fontWeight: 'bold', fontSize: '1.5em', }}>{selectedRecords.length}</span> 条记录 </div>}
          </div>

          <div style={{ marginTop: '16px', textAlign: 'center' }}>
            {(selectedRecords.length > 0) ? <Button onClick={(e)=> generateDocuments(selectedRecords, fields, selectedAttachmentField, primaryField)} variant="fill" color="primary" size="small" >导出 Word 文档</Button> : "请点击表格任意单元格"}
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
            <img src='https://s1.vika.cn/space/2021/12/29/5a4c225aed81490583cedbecf4bc3419' style={{ width: '30%' }} />
          </div>
          <div>请设置一个存储word模板的附件字段</div>

          <div style={{ marginTop: '16px', textAlign: 'center' }}>
            <Button onClick={() => openSettingArea()} variant="fill" color="primary" size="small" >前往设置</Button>
          </div>
        </div>
      }
    </div>
  );
};
