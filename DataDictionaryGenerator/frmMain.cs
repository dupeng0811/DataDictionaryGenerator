using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.Windows.Forms;
using System.IO;
using DataDictionaryGenerator.Model;
using RazorEngine;
using RazorEngine.Templating;

namespace DataDictionaryGenerator
{
    public partial class frmMain : Form
    {
        private readonly DataTable _dtInfo;
        public frmMain()
        {
            InitializeComponent();
            _dtInfo = new DataTable();
        }

        #region 自定义方法

        public ReturnMessage CheckCnnString(string cnnString)
        {
            ReturnMessage retMsg = new ReturnMessage(string.Empty, true);
            SqlConnection cnn = new SqlConnection(cnnString);
            try
            {
                cnn.Open();
            }
            catch (Exception ex)
            {
                retMsg.isSuccess = false;
                retMsg.Messages = ex.Message;
                return retMsg;
            }
            finally
            {
                if (cnn.State == ConnectionState.Open)
                {
                    cnn.Close();
                }
                cnn.Dispose();
            }
            return retMsg;
        }

        public ReturnMessage GetInfo(string cnnString)
        {
            ReturnMessage retMsg = new ReturnMessage(string.Empty, true);
            _dtInfo.Rows.Clear();
            string strQry = @"SELECT TOP (100) PERCENT d.name                  AS 表名,
                         CASE
                             WHEN a.colorder = 1 THEN isnull(f.value, '')
                             ELSE '' END                                       AS 表说明,
                         a.colorder                                            AS 字段序号,
                         a.name                                                AS 字段名,
                         CASE
                             WHEN COLUMNPROPERTY(a.id, a.name, 'IsIdentity')
                                 = 1 THEN '√'
                             ELSE '' END                                       AS 标识,
                         CASE
                             WHEN EXISTS
                                 (SELECT 1
                                  FROM dbo.sysindexes si
                                           INNER JOIN
                                       dbo.sysindexkeys sik ON si.id = sik.id AND si.indid = sik.indid
                                           INNER JOIN
                                       dbo.syscolumns sc ON sc.id = sik.id AND sc.colid = sik.colid
                                           INNER JOIN
                                       dbo.sysobjects so ON so.name = so.name AND so.xtype = 'PK'
                                  WHERE sc.id = a.id
                                    AND sc.colid = a.colid) THEN '√'
                             ELSE '' END                                       AS 主键,
                         b.name                                                AS 类型,
                         a.length                                              AS 长度,
                         COLUMNPROPERTY(a.id, a.name,
                                        'PRECISION')                           AS 精度,
                         ISNULL(COLUMNPROPERTY(a.id, a.name, 'Scale'), 0)      AS 小数位数,
                         CASE WHEN a.isnullable = 1 THEN '√' ELSE '' END       AS 允许空,
                         ISNULL(e.text, '')                                    AS 默认值,
                         ISNULL(g.value, '')                                   AS 字段说明,
                         d.crdate                                              AS 创建时间,
                         CASE WHEN a.colorder = 1 THEN d.refdate ELSE NULL END AS 更改时间
FROM dbo.syscolumns AS a
         LEFT OUTER JOIN
     dbo.systypes AS b ON a.xtype = b.xusertype
         INNER JOIN
     dbo.sysobjects AS d ON a.id = d.id AND d.xtype = 'U' AND d.status >= 0
         LEFT OUTER JOIN
     dbo.syscomments AS e ON a.cdefault = e.id
         LEFT OUTER JOIN
     sys.extended_properties AS g ON a.id = g.major_id AND a.colid = g.minor_id
         LEFT OUTER JOIN
     sys.extended_properties AS f ON d.id = f.major_id AND f.minor_id = 0
ORDER BY d.name, 字段序号";
            try
            {
                SqlDataAdapter da = new SqlDataAdapter(strQry, cnnString);
                da.Fill(_dtInfo);
                dgvData.DataSource = _dtInfo;
                return retMsg;
            }
            catch (Exception ex)
            {
                retMsg.isSuccess = false;
                retMsg.Messages = ex.Message;
                return retMsg;
            }
        }

        private ReturnMessage WriteMarkdown()
        {
            ReturnMessage retMsg = new ReturnMessage(string.Empty, true);
            try
            {
                Dictionary<string, List<TableColumnInfo>> templateDatas = new Dictionary<string, List<TableColumnInfo>>();
                int index = 1;
                foreach (DataRow dr in _dtInfo.Rows)
                {
                    string tableName = dr["表名"].ToString();
                    if (!templateDatas.ContainsKey(tableName))
                    {
                        templateDatas.Add(tableName,new List<TableColumnInfo>());
                    }
                    templateDatas[tableName].Add(new TableColumnInfo()
                    {
                        AllowNull = dr["允许空"].ToString(),
                        Description = dr["字段说明"].ToString(),
                        IsPrimary = dr["主键"].ToString(),
                        Name = dr["字段名"].ToString(),
                        Type = dr["类型"].ToString(),
                        TableName = tableName
                    });
                }
                string templatePath = Application.StartupPath + "\\Template\\template.cshtml";
                string templateContent = File.ReadAllText(templatePath, System.Text.Encoding.Default);
                string result = Engine.Razor.RunCompile(templateContent, "templateKey", null, templateDatas);
                
                string docAllPath = Application.StartupPath + "\\MarkdownDicFile.md";
                retMsg.Messages += docAllPath;

                if (File.Exists(docAllPath))
                {
                    File.Delete(docAllPath);
                }
                StreamWriter sw = File.AppendText(docAllPath); //保存到指定路径
                sw.Write(result);
                sw.Flush();
                sw.Close();

                return retMsg;
            }
            catch (Exception ex)
            {
                retMsg.isSuccess = false;
                retMsg.Messages = ex.Message;
                return retMsg;
            }
        }
        
        #endregion

        #region 委托方法

        private void btnBulid_Click(object sender, EventArgs e)
        {
            Cursor currCursor = this.Cursor;
            this.Cursor = Cursors.WaitCursor;

            ReturnMessage retMsg = CheckCnnString(txtCnnString.Text.Trim());
            if (!retMsg.isSuccess)
            {
                MessageBox.Show("数据库连接字符串错误，信息为：" + retMsg.Messages);
                this.Cursor = currCursor;
                return;
            }
            retMsg = GetInfo(txtCnnString.Text.Trim());
            if (!retMsg.isSuccess)
            {
                MessageBox.Show("读取数据库表结构错误，信息为：" + retMsg.Messages);
                this.Cursor = currCursor;
                return;
            }
            retMsg = WriteMarkdown();
            if (retMsg.isSuccess)
            {
                MessageBox.Show("文档生成成功!"+retMsg.Messages);
            }
            else
            {
                MessageBox.Show("文档生成失败，信息为：" + retMsg.Messages);
            }

            this.Cursor = currCursor;
        }

        #endregion
    }
}