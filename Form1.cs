using ActiveDs;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Messages;
using Microsoft.Xrm.Sdk.Query;
using RekTec.XStudio.CrmClient;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.OleDb;
using System.DirectoryServices;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace Rektec.Tools.UpdateUserRole
{
    public partial class Form1 : Form
    {
        IOrganizationService organizationServiceAdmin;

        public Form1()
        {
            InitializeComponent();
            organizationServiceAdmin = CrmServiceFactory.CreateOrganizationService(Guid.Empty);
        }

        private void button2_Click(object sender, EventArgs e)
        {
            OpenFileDialog dialog = new OpenFileDialog();
            dialog.Filter = "Excel files (*.xls;*.xlsx)|*.xls;*.xlsx|All files (*.*)|*.*";
            if (dialog.ShowDialog() == DialogResult.OK)
            {
                textBox1.Text = dialog.FileName;
            }
        }

        public DataSet GetAllDataTable(string fileFullPath)
        {
            DataSet dsResult = new DataSet();
            List<string> sheetNameList = GetAllSheetName(fileFullPath);
            string[] valideSheet = sheetNameList.Where((string a) => (a.Substring(a.Length - 2, 1) == "$" || a.Substring(a.Length - 1, 1) == "$") && !a.EndsWith("_")).ToArray();
            string[] array = valideSheet;
            foreach (string item in array)
            {
                string strSQL = $"select * from [{item}]";
                try
                {
                    DataTable dt = ExecuteDataTable(fileFullPath, strSQL);
                    DataTable newDT = dt.Copy();
                    newDT.TableName = item.Replace("'", "").TrimEnd('$');
                    dsResult.Tables.Add(newDT);
                }
                catch
                {
                }
            }
            return dsResult;
        }

        public List<string> GetAllSheetName(string fileFullPath)
        {
            List<string> result = new List<string>();
            OleDbConnection conn = new OleDbConnection(GetConnectString(fileFullPath));
            conn.Open();
            try
            {
                DataTable tables = conn.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, null);
                foreach (DataRow item in tables.Rows)
                {
                    result.Add(item["TABLE_NAME"].ToString());
                }
            }
            catch (Exception ex)
            {
                throw ex;
            }
            finally
            {
                conn.Close();
            }
            return result;
        }

        public DataTable ExecuteDataTable(string fileFullPath, string SQL)
        {
            OleDbConnection conn = new OleDbConnection(GetConnectString(fileFullPath));
            conn.Open();
            try
            {
                OleDbCommand command = new OleDbCommand(SQL, conn);
                OleDbDataAdapter adapter = new OleDbDataAdapter(command);
                DataSet ds = new DataSet();
                adapter.Fill(ds);
                return ds.Tables[0];
            }
            catch (Exception ex)
            {
                throw ex;
            }
            finally
            {
                conn.Close();
            }
        }

        public string GetConnectString(string fileFullPath)
        {
            string result = "";
            string extentsion = System.IO.Path.GetExtension(fileFullPath);
            string text = extentsion.ToUpper();
            if (!(text == ".XLS"))
            {
                if (text == ".XLSX")
                {
                    result = $"Provider=Microsoft.ACE.OLEDB.12.0;Data Source={fileFullPath};Extended Properties='Excel 12.0 Xml;HDR=YES;'";
                }
            }
            else
            {
                result = $"Provider=Microsoft.Jet.OLEDB.4.0;Data Source={fileFullPath};Extended Properties='Excel 8.0;IMEX=1;HDR=YES;'";
            }
            return result;
        }

        /// <summary>
        /// 导入角色
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void button1_Click(object sender, EventArgs e)
        {
            LoadingHelper.ShowLoadingScreen();//显示
            try
            {
                if (!File.Exists(textBox1.Text))
                {
                    ConcateLogMessage(this.richTextBox1, $"请选择可访问的文件，当前 {this.textBox1.Text}");

                    return;
                }

                if (string.IsNullOrWhiteSpace(this.comboBox1.Text))
                {
                    ConcateLogMessage(this.richTextBox1, $"请选择匹配系统用户的字段，当前 {this.comboBox1.Text}");

                    return;
                }

                DataSet ds = GetAllDataTable(textBox1.Text);
                IList<UserRoleItem> userRoles = new List<UserRoleItem>();

                foreach (DataRow row in ds.Tables[0].Rows)
                {
                    var rolename = string.Empty;
                    var username = string.Empty;

                    try
                    {
                        rolename = row["安全角色"].ToString().Trim();
                        username = row["用户"].ToString().Trim();

                        userRoles.Add(new UserRoleItem() { UserName = username, RoleName = rolename });
                    }
                    catch (Exception ex)
                    {
                        ConcateLogMessage(this.richTextBox1, $"文档值有问题，转文本失败，弄你");

                        return;
                    }
                }

                #region 移除用户当前配置角色
                if (this.checkBox1.Checked == true)
                {
                    var usergroups = userRoles.GroupBy(t => t.UserName).ToList();

                    foreach (var group in usergroups)
                    {
                        #region 查询目标系统用户
                        FetchExpression userexpression = new FetchExpression(string.Format($@"" +
                            " <fetch distinct='false' mapping='logical' > " +
                            "     <entity name='systemuser'> " +
                            "        <attribute name='systemuserid' />" +
                            "        <attribute name='businessunitid' />" +
                            "        <filter type='and'>" +
                            "            <condition attribute='{0}' operator='eq' value='{1}' />" +
                            "        </filter>" +
                            "     </entity> " +
                            " </fetch>", this.comboBox1.Text.Trim(), group.Key));

                        EntityCollection userentityList = organizationServiceAdmin.RetrieveMultiple(userexpression);

                        if (userentityList == null || userentityList.Entities.Count <= 0)
                        {
                            ConcateLogMessage(this.richTextBox1, $"CRM系统中没有找到用户 {group.Key}");

                            continue;
                        }

                        if (userentityList.Entities.Count > 1)
                        {
                            ConcateLogMessage(this.richTextBox1, $"CRM系统中存在多个同名用户 {group.Key}");

                            continue;
                        }
                        #endregion 查询目标系统用户

                        var systemuserid = userentityList.Entities[0].GetAttributeValue<Guid>("systemuserid");
                        var businessunitid = userentityList.Entities[0].GetAttributeValue<EntityReference>("businessunitid");

                        #region 查询用户已配置角色
                        string userrolesexpression = $@"<fetch distinct='false' mapping='logical'>" +
                             " <entity name='systemuserroles'>" +
                             "   <attribute name='roleid' />" +
                             "   <link-entity name='role' from='roleid' to='roleid'>" +
                            "      <attribute name='name' alias='rolenmae' />" +
                            "      <attribute name='businessunitid' alias='businessunitid' />" +
                            "    </link-entity>" +
                             "   <filter type='and'>" +
                            "      <condition entityname='systemuserroles' attribute='systemuserid' operator='eq' value='{0}' />" +
                             "   </filter>" +
                            "  </entity>" +
                            "</fetch>";

                        EntityCollection userroleentityList = organizationServiceAdmin.RetrieveMultiple(new FetchExpression(string.Format(userrolesexpression, systemuserid.ToString())));
                        #endregion 查询用户已配置角色

                        if (userroleentityList != null && userroleentityList.Entities != null && userroleentityList.Entities.Count > 0)
                        {
                            foreach (var userrole in userroleentityList.Entities)
                            {
                                var roleid = userrole.GetAttributeValue<Guid>("roleid");

                                DisassociateRequest disassociateRequest = new DisassociateRequest
                                {
                                    Target = new EntityReference("systemuser", systemuserid),
                                    RelatedEntities = new EntityReferenceCollection
                            {
                                new EntityReference("role", roleid)
                            },
                                    Relationship = new Relationship("systemuserroles_association")
                                };

                                organizationServiceAdmin.Execute(disassociateRequest);
                            }
                        }
                    }
                }
                #endregion 移除用户当前配置角色

                foreach (var item in userRoles)
                {
                    var rolename = item.RoleName;
                    var username = item.UserName;

                    #region 查询目标系统用户
                    FetchExpression userexpression = new FetchExpression(string.Format($@"" +
                        " <fetch distinct='false' mapping='logical' > " +
                        "     <entity name='systemuser'> " +
                        "        <attribute name='systemuserid' />" +
                        "        <attribute name='businessunitid' />" +
                        "        <filter type='and'>" +
                        "            <condition attribute='{0}' operator='eq' value='{1}' />" +
                        "        </filter>" +
                        "     </entity> " +
                        " </fetch>", this.comboBox1.Text.Trim(), username));

                    EntityCollection userentityList = organizationServiceAdmin.RetrieveMultiple(userexpression);

                    if (userentityList == null || userentityList.Entities.Count <= 0)
                    {
                        ConcateLogMessage(this.richTextBox1, $"CRM系统中没有找到用户 {username}");

                        continue;
                    }

                    if (userentityList.Entities.Count > 1)
                    {
                        ConcateLogMessage(this.richTextBox1, $"CRM系统中存在多个同名用户 {username}");

                        continue;
                    }
                    #endregion 查询目标系统用户

                    var systemuserid = userentityList.Entities[0].GetAttributeValue<Guid>("systemuserid");
                    var businessunitid = userentityList.Entities[0].GetAttributeValue<EntityReference>("businessunitid");

                    #region 查询目标角色
                    string roleexpression = $@" <fetch distinct='false' mapping='logical' > " +
                        "     <entity name='role'> " +
                        "        <attribute name='roleid' />" +
                        "        <filter type='and'>" +
                        "             <condition attribute='name' operator='eq' value='{0}' />" +
                         "             <condition attribute='businessunitid' operator='eq' value='{1}' />" +
                        "        </filter>" +
                        "     </entity> " +
                        " </fetch>";


                    EntityCollection roleentityList = organizationServiceAdmin.RetrieveMultiple(new FetchExpression(string.Format(roleexpression, rolename, businessunitid.Id.ToString())));


                    if (roleentityList == null || roleentityList.Entities.Count <= 0)
                    {
                        ConcateLogMessage(this.richTextBox1, $"CRM系统中没有找到安全角色 {rolename},{businessunitid.Name}");

                        continue;
                    }

                    if (roleentityList.Entities.Count > 1)
                    {
                        ConcateLogMessage(this.richTextBox1, $"CRM系统中存在多个同名安全角色 {rolename},{businessunitid.Name}");

                        continue;
                    }
                    #endregion 查询目标角色

                    var roleid = roleentityList.Entities[0].GetAttributeValue<Guid>("roleid");

                    #region 查询是否已配置目标角色
                    string userrolesexpression = $@"<fetch distinct='false' mapping='logical'>" +
                         " <entity name='systemuserroles'>" +
                         "   <attribute name='roleid' />" +
                         "   <link-entity name='role' from='roleid' to='roleid'>" +
                        "      <attribute name='name' alias='rolenmae' />" +
                        "      <attribute name='businessunitid' alias='businessunitid' />" +
                        "    </link-entity>" +
                         "   <filter type='and'>" +
                        "      <condition entityname='systemuserroles' attribute='systemuserid' operator='eq' value='{0}' />" +
                        "     <condition entityname='role' attribute='roleid' operator='eq' value='{1}' />" +
                         "   </filter>" +
                        "  </entity>" +
                        "</fetch>";

                    EntityCollection userroleentityList = organizationServiceAdmin.RetrieveMultiple(new FetchExpression(string.Format(userrolesexpression, systemuserid.ToString(), roleid.ToString())));
                    #endregion 查询是否已配置目标角色

                    #region 添加/移除角色
                    if (userroleentityList != null && userroleentityList.Entities != null && userroleentityList.Entities.Count > 0)
                    {
                        if (this.checkBox2.Checked == true)
                        {
                            DisassociateRequest disassociateRequest = new DisassociateRequest
                            {
                                Target = new EntityReference("systemuser", systemuserid),
                                RelatedEntities = new EntityReferenceCollection
                            {
                                new EntityReference("role", roleid)
                            },
                                Relationship = new Relationship("systemuserroles_association")
                            };

                            organizationServiceAdmin.Execute(disassociateRequest);

                            if (this.checkBox6.Checked == false)
                            {
                                ConcateLogMessage(this.richTextBox1, $"移除用户角色成功 {username},{rolename}");
                            }
                        }
                        else
                        {
                            if (this.checkBox6.Checked == false)
                            {
                                ConcateLogMessage(this.richTextBox1, $"用户角色已存在，跳过 {username},{rolename}");
                            }
                        }

                        continue;
                    }
                    else
                    {
                        if (this.checkBox2.Checked == true)
                        {
                            if (this.checkBox6.Checked == false)
                            {
                                ConcateLogMessage(this.richTextBox1, $"用户当前没有角色 {username},{rolename},无需移除");
                            }

                            continue;
                        }
                    }

                    AssociateRequest associateRequest = new AssociateRequest
                    {
                        Target = new EntityReference("systemuser", systemuserid),
                        RelatedEntities = new EntityReferenceCollection
                                {
                                    new EntityReference("role", roleid)
                                },
                        Relationship = new Relationship("systemuserroles_association")
                    };

                    organizationServiceAdmin.Execute(associateRequest);

                    if (this.checkBox6.Checked == false)
                    {
                        ConcateLogMessage(this.richTextBox1, $"添加用户角色成功 {username},{rolename}");
                    }
                    #endregion 添加/移除角色
                }

                ConcateLogMessage(this.richTextBox1, $"导入完毕");

                //LoadingHelper.CloseForm();//关闭
            }
            catch (Exception ex)
            {
                ConcateLogMessage(this.richTextBox1, $"导入失败");
            }
            finally
            {
                LoadingHelper.CloseForm();//关闭
            }
        }

        private void checkBox1_CheckedChanged(object sender, EventArgs e)
        {
            if ((sender as CheckBox).Checked == true)
            {
                foreach (CheckBox chk in (sender as CheckBox).Parent.Controls)
                {
                    if (chk != sender)
                    {
                        chk.Checked = false;
                    }
                }
            }
        }

        /// <summary>
        /// 更新部门
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void button3_Click(object sender, EventArgs e)
        {
            LoadingHelper.ShowLoadingScreen();//显示
            try
            {
                if (!File.Exists(textBox1.Text))
                {
                    ConcateLogMessage(this.richTextBox1, $"请选择可访问的文件，当前 {this.textBox1.Text}");

                    return;
                }

                if (string.IsNullOrWhiteSpace(this.comboBox1.Text))
                {
                    ConcateLogMessage(this.richTextBox1, $"请选择匹配系统用户的字段，当前 {this.comboBox1.Text}");

                    return;
                }

                DataSet ds = GetAllDataTable(textBox1.Text);
                IList<UserBusinessunitItem> userBus = new List<UserBusinessunitItem>();

                foreach (DataRow row in ds.Tables[0].Rows)
                {
                    var bucondition = string.Empty;
                    var buname = string.Empty;
                    var username = string.Empty;

                    try
                    {
                        bucondition = row["业务部门匹配值"].ToString().Trim();
                        buname = row["业务部门"].ToString().Trim();
                        username = row["用户"].ToString().Trim();

                        userBus.Add(new UserBusinessunitItem() { UserName = username, BuCondition = bucondition, BuName = buname });
                    }
                    catch (Exception ex)
                    {
                        ConcateLogMessage(this.richTextBox1, $"文档值有问题，转文本失败，弄你");

                        return;
                    }
                }

                foreach (var item in userBus)
                {
                    var bucondition = item.BuCondition;
                    var username = item.UserName;
                    var buname = item.BuName;

                    #region 查询目标系统用户
                    FetchExpression userexpression = new FetchExpression(string.Format($@"" +
                        " <fetch distinct='false' mapping='logical' > " +
                        "     <entity name='systemuser'> " +
                        "        <attribute name='systemuserid' />" +
                        "        <attribute name='businessunitid' />" +
                        "        <filter type='and'>" +
                        "            <condition attribute='{0}' operator='eq' value='{1}' />" +
                        "        </filter>" +
                        "     </entity> " +
                        " </fetch>", this.comboBox1.Text.Trim(), username));

                    EntityCollection userentityList = organizationServiceAdmin.RetrieveMultiple(userexpression);

                    if (userentityList == null || userentityList.Entities.Count <= 0)
                    {
                        ConcateLogMessage(this.richTextBox1, $"CRM系统中没有找到用户 {username}");

                        continue;
                    }

                    if (userentityList.Entities.Count > 1)
                    {
                        ConcateLogMessage(this.richTextBox1, $"CRM系统中存在多个同名用户 {username}");

                        continue;
                    }
                    #endregion 查询目标系统用户

                    var systemuserid = userentityList.Entities[0].GetAttributeValue<Guid>("systemuserid");
                    var businessunitid = userentityList.Entities[0].GetAttributeValue<EntityReference>("businessunitid");

                    #region 查询目标部门
                    string roleexpression = $@" <fetch distinct='false' mapping='logical' > " +
                        "     <entity name='businessunit'> " +
                        "        <attribute name='businessunitid' />" +
                        "        <filter type='and'>" +
                        "             <condition attribute='{0}' operator='eq' value='{1}' />" +
                         "             <condition attribute='isdisabled' operator='eq' value='0' />" +
                        "        </filter>" +
                        "     </entity> " +
                        " </fetch>";

                    EntityCollection buentityList = organizationServiceAdmin.RetrieveMultiple(new FetchExpression(string.Format(roleexpression, this.comboBox2.Text.Trim(), bucondition)));

                    if (buentityList == null || buentityList.Entities.Count <= 0)
                    {
                        ConcateLogMessage(this.richTextBox1, $"CRM系统中没有找到业务部门：业务部门匹配值 {bucondition}");

                        continue;
                    }

                    if (buentityList.Entities.Count > 1)
                    {
                        ConcateLogMessage(this.richTextBox1, $"CRM系统中存在重复业务部门：业务部门匹配值 {bucondition}");

                        continue;
                    }
                    #endregion 查询目标部门

                    var buid = buentityList.Entities[0].GetAttributeValue<Guid>("businessunitid");

                    #region 更新业务部门
                    Entity user = new Entity("systemuser");
                    user.Id = systemuserid;

                    user["businessunitid"] = new EntityReference("businessunit", buid);

                    organizationServiceAdmin.Update(user);

                    ConcateLogMessage(this.richTextBox1, $"业务部门更新成功 {username},{bucondition}");
                    #endregion 更新业务部门
                }

                ConcateLogMessage(this.richTextBox1, $"业务部门更新完毕");
            }
            catch (Exception ex)
            {
                ConcateLogMessage(this.richTextBox1, $"业务部门更新失败");
            }
            finally
            {
                LoadingHelper.CloseForm();//关闭
            }
        }

        void ConcateLogMessage(RichTextBox richTextBox, string message)
        {
            richTextBox.Text = message + Environment.NewLine + richTextBox.Text;
        }

        private void button4_Click(object sender, EventArgs e)
        {
            IsConnected();
        }

        private void IsConnected()
        {
            try
            {
                DirectoryEntry domain = new DirectoryEntry();
                domain.Path = this.textBox2.Text;
                domain.Username = this.textBox5.Text;
                domain.Password = this.textBox6.Text;
                domain.AuthenticationType = AuthenticationTypes.Secure;
                domain.RefreshCache();


                ConcateLogMessage(this.richTextBox4, $"AD服务器连接成功");
            }

            catch (Exception ex)
            {
                ConcateLogMessage(this.richTextBox4, $"AD服务器连接失败" + ex.Message);
            }
        }
        private void button5_Click(object sender, EventArgs e)
        {
            LoadingHelper.ShowLoadingScreen();//显示

            try
            {
                if (!File.Exists(textBox4.Text))
                {
                    ConcateLogMessage(this.richTextBox4, $"请选择可访问的文件，当前 {this.textBox4.Text}");

                    return;
                }

                if (string.IsNullOrWhiteSpace(textBox5.Text))
                {
                    ConcateLogMessage(this.richTextBox4, $"域控管理员账号（crmadmin）不要提供吗？，当前 {this.textBox5.Text}");

                    return;
                }


                if (string.IsNullOrWhiteSpace(textBox6.Text))
                {
                    ConcateLogMessage(this.richTextBox4, $"域控管理员密码（crmadmin）不要提供吗？，当前 {this.textBox6.Text}");

                    return;
                }

                DataSet ds = GetAllDataTable(textBox4.Text);
                IList<UserItem> adusers = new List<UserItem>();

                foreach (DataRow row in ds.Tables[0].Rows)
                {
                    var loginname = string.Empty;
                    var password = string.Empty;
                    var username = string.Empty;
                    var email = string.Empty;
                    var accountoption = string.Empty;

                    try
                    {
                        username = row["姓名"].ToString().Trim();
                        loginname = row["用户登录名"].ToString().Trim();
                        password = row["密码"].ToString().Trim();
                        email = row["邮箱"].ToString().Trim();
                        accountoption = row["账户选项"].ToString().Trim();

                        if (string.IsNullOrWhiteSpace(password))
                        {
                            password = "P@ssw0rd";
                        }

                        var user = new UserItem() { UserName = username, UserLoginName = loginname, Password = password, Email = email, AccountOption = accountoption };

                        if (this.checkBox4.Checked == true)
                        {
                            if (string.IsNullOrWhiteSpace(this.textBox7.Text))
                            {
                                ConcateLogMessage(this.richTextBox4, $"自定义账户选项值提供一下啊，弄你");

                                return;
                            }

                            Int32 accountOptionValue;

                            if (Int32.TryParse(this.textBox7.Text, out accountOptionValue) == false)
                            {
                                ConcateLogMessage(this.richTextBox4, $"自定义账户选项值是整数类型，网上查一下，弄你");

                                return;
                            }

                            user.AccountOptionValue = Int32.Parse(this.textBox7.Text.Trim());
                        }
                        else
                        {
                            switch (user.AccountOption)
                            {
                                case "用户下次登陆时须更改密码":
                                    user.AccountOptionValue = 544;
                                    break;
                                case "用户不能更改密码":
                                    user.AccountOptionValue = 576;
                                    break;
                                case "密码永不过期":
                                    user.AccountOptionValue = 66048;
                                    break;
                                case "账户已禁用":
                                    user.AccountOptionValue = 514;
                                    break;
                                case "正常":
                                    user.AccountOptionValue = 512;
                                    break;
                                default:
                                    user.AccountOption = "正常";
                                    user.AccountOptionValue = 512;
                                    break;
                            }
                        }

                        adusers.Add(user);
                    }
                    catch (Exception ex)
                    {
                        ConcateLogMessage(this.richTextBox4, $"文档值有问题，转文本失败，弄你");

                        return;
                    }
                }

                foreach (var item in adusers)
                {
                    AddUser2AD(item);
                }
            }
            catch (Exception ex)
            {
                ConcateLogMessage(this.richTextBox4, $"新增AD用户失败");
            }
            finally
            {
                LoadingHelper.CloseForm();//关闭
            }
        }

        public void AddUser2AD(UserItem user)
        {
            string strname = "CN=" + user.UserName;
            try
            {
                // strADAccount ，strADPassword为AD管理员账户和密码
                DirectoryEntry objDE = new DirectoryEntry(this.textBox2.Text + this.textBox3.Text, this.textBox5.Text, this.textBox6.Text);

                DirectoryEntries objDES = objDE.Children;
                DirectoryEntry myDE = objDES.Add(strname, "User");
                myDE.Properties["userPrincipalName"].Value = user.UserLoginName;
                myDE.Properties["name"].Value = user.UserName;
                myDE.Properties["sAMAccountName"].Value = user.UserLoginName;
                //myDE.Properties["userWorkstations"].Value = "JTFWPTQASAD01";
                if (user.AccountOption == "用户下次登陆时须更改密码" || user.AccountOption == "账户已禁用")
                {
                    myDE.Properties["UserPassword"].Add(user.Password);
                    myDE.Properties["userAccountControl"].Value = user.AccountOptionValue;
                }

                if (!string.IsNullOrWhiteSpace(user.Email))
                {
                    myDE.Properties["mail"].Value = user.Email;
                }

                myDE.CommitChanges();

                if (user.AccountOption == "用户不能更改密码" || user.AccountOption == "密码永不过期" || user.AccountOption == "正常")
                {
                    //设置密码 ，这里需要到服务器执行
                    //（需要引用COM: Active DS Type Library,引用命名空间using ActiveDs;）
                    ActiveDs.IADsUser objUser = myDE.NativeObject as IADsUser;
                    objUser.SetPassword(user.Password);
                    //设置用户状态：密码永不过期（65536）+用户正常（512）= 66048
                    objUser.Put("userAccountControl", user.AccountOptionValue);

                    objUser.SetInfo();
                }

                if (this.checkBox5.Checked == false)
                {
                    ConcateLogMessage(this.richTextBox4, $"AD服务器新增用户成功" + user.UserName + "[" + user.UserLoginName + "]");
                }
            }
            catch (Exception ex)
            {
                ConcateLogMessage(this.richTextBox4, $"AD服务器新增用户失败：" + ex.Message);
            }
        }

        public void ps(ListBox box, string s)
        {
            String line = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss") + " " + s;
            box.Items.Add(line);

        }

        private void button6_Click(object sender, EventArgs e)
        {
            OpenFileDialog dialog = new OpenFileDialog();
            dialog.Filter = "Excel files (*.xls;*.xlsx)|*.xls;*.xlsx|All files (*.*)|*.*";
            if (dialog.ShowDialog() == DialogResult.OK)
            {
                textBox4.Text = dialog.FileName;
            }
        }

        private void button7_Click(object sender, EventArgs e)
        {
            LoadingHelper.ShowLoadingScreen();//显示

            try
            {
                if (!File.Exists(textBox4.Text))
                {
                    ConcateLogMessage(this.richTextBox4, $"请选择可访问的文件，当前 {this.textBox4.Text}");

                    return;
                }

                DataSet ds = GetAllDataTable(textBox4.Text);
                IList<UserItem> crmusers = new List<UserItem>();

                foreach (DataRow row in ds.Tables[0].Rows)
                {
                    var loginname = string.Empty;
                    var password = string.Empty;
                    var username = string.Empty;
                    var email = string.Empty;
                    var accountoption = string.Empty;

                    try
                    {
                        var user = new UserItem()
                        {
                            DomainName = row["用户名"].ToString().Trim(),
                            LastName = row["姓"].ToString().Trim(),
                            FirstName = row["名"].ToString().Trim(),
                            BuCondition = row["业务部门匹配值"].ToString().Trim(),
                            BuName = row["业务部门"].ToString().Trim(),
                            Internalemailaddress = row["主要电子邮件"].ToString().Trim(),
                            Mobilephone = row["移动电话"].ToString().Trim(),
                            Address1Telephone1 = row["主要电话"].ToString().Trim(),
                            Jobnumber = row["工号"].ToString().Trim(),
                            Jobtitle = row["职务"].ToString().Trim()
                        };

                        crmusers.Add(user);
                    }
                    catch (Exception ex)
                    {
                        ConcateLogMessage(this.richTextBox4, $"文档值有问题，转文本失败，弄你");

                        return;
                    }
                }

                foreach (var user in crmusers)
                {
                    try
                    {
                        Entity systemuser = new Entity("systemuser");

                        systemuser["domainname"] = user.DomainName;
                        systemuser["lastname"] = user.LastName;
                        systemuser["firstname"] = user.FirstName;

                        if (!string.IsNullOrWhiteSpace(user.Internalemailaddress))
                        {
                            systemuser["internalemailaddress"] = user.Internalemailaddress;
                        }

                        if (!string.IsNullOrWhiteSpace(user.Mobilephone))
                        {
                            systemuser["mobilephone"] = user.Mobilephone;
                        }

                        if (!string.IsNullOrWhiteSpace(user.Address1Telephone1))
                        {
                            systemuser["address1_telephone1"] = user.Address1Telephone1;
                        }

                        if (!string.IsNullOrWhiteSpace(user.Jobnumber))
                        {
                            systemuser["new_jobnumber"] = user.Jobnumber;
                        }

                        if (!string.IsNullOrWhiteSpace(user.Jobtitle))
                        {
                            systemuser["jobtitle"] = user.Jobtitle;
                        }

                        #region 查询目标部门
                        string roleexpression = $@" <fetch distinct='false' mapping='logical' > " +
                            "     <entity name='businessunit'> " +
                            "        <attribute name='businessunitid' />" +
                            "        <filter type='and'>" +
                            "             <condition attribute='{0}' operator='eq' value='{1}' />" +
                             "             <condition attribute='isdisabled' operator='eq' value='0' />" +
                            "        </filter>" +
                            "     </entity> " +
                            " </fetch>";

                        EntityCollection buentityList = organizationServiceAdmin.RetrieveMultiple(new FetchExpression(string.Format(roleexpression, this.comboBox2.Text.Trim(), user.BuCondition)));

                        if (buentityList == null || buentityList.Entities.Count <= 0)
                        {
                            ConcateLogMessage(this.richTextBox4, $"CRM系统中没有找到业务部门：业务部门匹配值 {user.BuCondition}");

                            continue;
                        }

                        if (buentityList.Entities.Count > 1)
                        {
                            ConcateLogMessage(this.richTextBox4, $"CRM系统中存在重复业务部门：业务部门匹配值 {user.BuCondition}");

                            continue;
                        }
                        #endregion 查询目标部门

                        var buid = buentityList.Entities[0].GetAttributeValue<Guid>("businessunitid");

                        #region 查询目标系统用户
                        FetchExpression userexpression = new FetchExpression(string.Format($@"" +
                            " <fetch distinct='false' mapping='logical' > " +
                            "     <entity name='systemuser'> " +
                            "        <attribute name='systemuserid' />" +
                            "        <attribute name='businessunitid' />" +
                            "        <filter type='and'>" +
                            "            <condition attribute='{0}' operator='eq' value='{1}' />" +
                            "        </filter>" +
                            "     </entity> " +
                            " </fetch>", this.comboBox4.Text.Trim(), user.DomainName));

                        EntityCollection userentityList = organizationServiceAdmin.RetrieveMultiple(userexpression);

                        if (userentityList.Entities.Count > 1)
                        {
                            ConcateLogMessage(this.richTextBox4, $"CRM系统中存在多个同名用户 {user.DomainName}");

                            continue;
                        }
                        #endregion 查询目标系统用户

                        if (userentityList.Entities.Count == 1)
                        {
                            var systemuserid = userentityList.Entities[0].GetAttributeValue<Guid>("systemuserid");
                            var businessunitid = userentityList.Entities[0].GetAttributeValue<EntityReference>("businessunitid");

                            systemuser.Id = systemuserid;

                            if (buid != businessunitid.Id)
                            {
                                systemuser["businessunitid"] = new EntityReference("businessunit", buid);
                            }

                            organizationServiceAdmin.Update(systemuser);

                            if (this.checkBox5.Checked == false)
                            {
                                ConcateLogMessage(this.richTextBox4, $"更新CRM用户成功{user.DomainName}");
                            }
                        }
                        else
                        {
                            systemuser.Id = Guid.NewGuid();

                            systemuser["businessunitid"] = new EntityReference("businessunit", buid);

                            organizationServiceAdmin.Create(systemuser);

                            if (this.checkBox5.Checked == false)
                            {
                                ConcateLogMessage(this.richTextBox4, $"新增CRM用户成功{user.DomainName}");
                            }
                        }
                    }
                    catch (Exception ex)
                    {
                        ConcateLogMessage(this.richTextBox4, $"新增/修改CRM用户失败：{user.DomainName}" + ex.Message);
                    }
                }
            }
            catch (Exception ex)
            {
                ConcateLogMessage(this.richTextBox4, $"新增/修改CRM用户失败：" + ex.Message);
            }
            finally
            {
                LoadingHelper.CloseForm();//关闭
            }
        }

        private void button7_Click_1(object sender, EventArgs e)
        {
            this.richTextBox4.Text = "";
        }

        private void button9_Click(object sender, EventArgs e)
        {
            this.richTextBox1.Text = "";
        }

        private void button10_Click(object sender, EventArgs e)
        {
            this.button11_Click(sender, e);

            if (string.IsNullOrWhiteSpace(this.textBox2.Text))
            {
                ConcateLogMessage(this.richTextBox4, $"请点击'获取LDAP域'");

                return;
            }

            this.treeView1.Nodes.Clear();

            DirectoryEntry rootEntry = new DirectoryEntry(this.textBox2.Text);

            DirectorySearcher dsFindOUs = new DirectorySearcher(rootEntry);

            dsFindOUs.Filter = "(|(&(objectClass=organizationalUnit)(!ou=Domain Controllers))(&(objectClass=container)(cn=Users)))";
            //dsFindOUs.Filter = "(|(objectClass=organizationalUnit)(objectClass=container))";
            //(| (objectClass = organizationalUnit)(objectCategory = group)(objectCategory = computer)(objectClass = domainDNS))

            dsFindOUs.SearchScope = SearchScope.Subtree;

            dsFindOUs.PropertiesToLoad.Add("displayName");

            var findOus = dsFindOUs.FindAll();

            foreach (SearchResult result in findOus)
            {
                //ConcateLogMessage(this.richTextBox4, RekTec.Crm.Common.Helper.JsonHelper.Serialize(result));
                //ConcateLogMessage(this.richTextBox4, result.Path);
                treeView1.Nodes.Add(new TreeNode(result.Path));
            }

            if (findOus != null && findOus.Count > 0)
            {
                this.textBox3.Text = findOus[0].Path;
            }
        }

        private void button11_Click(object sender, EventArgs e)
        {
            DirectoryEntry rootEntry = new DirectoryEntry("LDAP://localhost");

            DirectorySearcher dsFindOUs = new DirectorySearcher(rootEntry);

            dsFindOUs.Filter = "(objectClass=domainDNS)";
            //(| (objectClass = organizationalUnit)(objectCategory = group)(objectCategory = computer)(objectClass = domainDNS))

            dsFindOUs.SearchScope = SearchScope.Subtree;

            dsFindOUs.PropertiesToLoad.Add("displayName");

            foreach (SearchResult result in dsFindOUs.FindAll())
            {
                //ConcateLogMessage(this.richTextBox4, result.Path);
                this.textBox2.Text = result.Path;
            }
        }

        private void treeView1_NodeMouseClick(object sender, TreeNodeMouseClickEventArgs e)
        {
            this.textBox3.Text = e.Node.Text;
        }
    }

    public class UserRoleItem
    {
        public string UserName { get; set; }
        public string RoleName { get; set; }
    }

    public class UserBusinessunitItem
    {
        public string UserName { get; set; }
        public string BuCondition { get; set; }
        public string BuName { get; set; }
    }

    public class UserItem
    {
        public string UserName { get; set; }
        public string UserLoginName { get; set; }
        public string Email { get; set; }
        public string Password { get; set; }
        public string AccountOption { get; set; }
        public Int32 AccountOptionValue { get; set; }

        #region CRM用户
        public string DomainName { get; set; }
        public string LastName { get; set; }
        public string FirstName { get; set; }
        public string BuCondition { get; set; }
        public string BuName { get; set; }
        public string Internalemailaddress { get; set; }
        public string Mobilephone { get; set; }
        public string Address1Telephone1 { get; set; }
        public string Jobnumber { get; set; }
        public string Jobtitle { get; set; }

        #endregion CRM用户
    }
}
