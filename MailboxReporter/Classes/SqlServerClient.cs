using System;
using System.Configuration;
using System.Data;
using System.Threading.Tasks;
using Microsoft.Data.SqlClient;

namespace MailboxReporter.Classes
{
    public class SqlServerClient : IDisposable
    {
        private readonly SqlConnection _reportConnection =
            new SqlConnection(ConfigurationManager.ConnectionStrings["MailboxReporterConnection"].ToString());

        public void Dispose()
        {
            if (IsOpen())
                Disconnect();

            _reportConnection?.Dispose();
        }

        public async Task Connect()
        {
            await _reportConnection.OpenAsync();
        }

        public void Disconnect()
        {
            _reportConnection.Close();
        }

        private bool IsOpen()
        {
            return _reportConnection != null && _reportConnection.State > ConnectionState.Closed &&
                   !_reportConnection.State.Equals(ConnectionState.Broken);
        }

        public async Task AddOrUpdate(string mailboxAddress, EmailRecord email)
        {
            if (!IsOpen())
                await Connect();

            var sqlCmd = new SqlCommand("MailboxReporter_AddOrUpdate", _reportConnection)
            {
                CommandType = CommandType.StoredProcedure,
                Parameters =
                {
                    new SqlParameter
                    {
                        ParameterName = "@Id", SqlDbType = SqlDbType.VarChar,
                        Value = string.IsNullOrEmpty(email.Id) ? "" : email.Id
                    },
                    new SqlParameter
                        {ParameterName = "@MailboxAddress", SqlDbType = SqlDbType.VarChar, Value = mailboxAddress},
                    new SqlParameter
                    {
                        ParameterName = "@InternetMessageId", SqlDbType = SqlDbType.VarChar,
                        Value = email.InternetMessageId
                    },
                    new SqlParameter
                    {
                        ParameterName = "@ConversationId", SqlDbType = SqlDbType.VarChar,
                        Value = string.IsNullOrEmpty(email.ConversationId) ? "" : email.ConversationId
                    },
                    new SqlParameter
                        {ParameterName = "@SentDate", SqlDbType = SqlDbType.DateTime, Value = email.SentDate},
                    new SqlParameter
                        {ParameterName = "@ReceivedDate", SqlDbType = SqlDbType.DateTime, Value = email.ReceivedDate},
                    new SqlParameter
                        {ParameterName = "@CreatedDate", SqlDbType = SqlDbType.DateTime, Value = email.CreatedDate},
                    new SqlParameter
                        {ParameterName = "@ModifiedDate", SqlDbType = SqlDbType.DateTime, Value = email.ModifiedDate},
                    new SqlParameter
                    {
                        ParameterName = "@ModifiedName", SqlDbType = SqlDbType.NVarChar,
                        Value = string.IsNullOrEmpty(email.ModifiedName) ? "" : email.ModifiedName
                    },
                    new SqlParameter
                    {
                        ParameterName = "@FromName", SqlDbType = SqlDbType.NVarChar,
                        Value = string.IsNullOrEmpty(email.FromName) ? "" : email.FromName
                    },
                    new SqlParameter
                    {
                        ParameterName = "@FromAddress", SqlDbType = SqlDbType.NVarChar,
                        Value = string.IsNullOrEmpty(email.FromAddress) ? "" : email.FromAddress
                    },
                    new SqlParameter
                    {
                        ParameterName = "@ReplyToName", SqlDbType = SqlDbType.NVarChar,
                        Value = string.IsNullOrEmpty(email.ReplyToName) ? "" : email.ReplyToName
                    },
                    new SqlParameter
                    {
                        ParameterName = "@ReplyToAddress", SqlDbType = SqlDbType.NVarChar,
                        Value = string.IsNullOrEmpty(email.ReplyToAddress) ? "" : email.ReplyToAddress
                    },
                    new SqlParameter
                    {
                        ParameterName = "@ToName", SqlDbType = SqlDbType.NVarChar,
                        Value = string.IsNullOrEmpty(email.ToName) ? "" : email.ToName
                    },
                    new SqlParameter
                    {
                        ParameterName = "@ToAddress", SqlDbType = SqlDbType.NVarChar,
                        Value = string.IsNullOrEmpty(email.ToAddress) ? "" : email.ToAddress
                    },
                    new SqlParameter
                    {
                        ParameterName = "@CcName", SqlDbType = SqlDbType.NVarChar,
                        Value = string.IsNullOrEmpty(email.CcName) ? "" : email.CcName
                    },
                    new SqlParameter
                    {
                        ParameterName = "@CcAddress", SqlDbType = SqlDbType.NVarChar,
                        Value = string.IsNullOrEmpty(email.CcAddress) ? "" : email.CcAddress
                    },
                    new SqlParameter
                    {
                        ParameterName = "@Priority", SqlDbType = SqlDbType.NVarChar, Value = email.Priority.ToString()
                    },
                    new SqlParameter
                    {
                        ParameterName = "@Subject", SqlDbType = SqlDbType.NVarChar,
                        Value = string.IsNullOrEmpty(email.Subject) ? "" : email.Subject
                    },
                    new SqlParameter
                    {
                        ParameterName = "@Body", SqlDbType = SqlDbType.NVarChar,
                        Value = string.IsNullOrEmpty(email.Body) ? "" : email.Body
                    },
                    new SqlParameter
                        {ParameterName = "@BodyType", SqlDbType = SqlDbType.NVarChar, Value = email.BodyType},
                    new SqlParameter {ParameterName = "@Size", SqlDbType = SqlDbType.Int, Value = email.Size},
                    new SqlParameter
                        {ParameterName = "@AttachmentCount", SqlDbType = SqlDbType.Int, Value = email.AttachmentCount},
                    new SqlParameter
                        {ParameterName = "@Attachments", SqlDbType = SqlDbType.NVarChar, Value = email.Attachments},
                    new SqlParameter {ParameterName = "@IsRead", SqlDbType = SqlDbType.Bit, Value = email.IsRead}
                }
            };

            await sqlCmd.ExecuteNonQueryAsync();
            sqlCmd.Dispose();
        }
    }
}