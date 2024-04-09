using Google.Apis.Auth.OAuth2;
using Google.Apis.Gmail.v1;
using Google.Apis.Gmail.v1.Data;
using Google.Apis.Services;
using Google.Apis.Util.Store;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.Data.SqlClient;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading;
using Dapper;

namespace GmailAPI.APIHelper
{
    public class GmailHelper
    {
        static string[] Scopes = { GmailService.Scope.MailGoogleCom };
        static string ApplicationName = "Gmail API Application";

        public static GmailService GetService()
        {
            UserCredential credential;
            using (FileStream stream = new FileStream(Convert.ToString(ConfigurationManager.AppSettings["ClientInfo"]), FileMode.Open, FileAccess.Read))
            {
                String FolderPath = Convert.ToString(ConfigurationManager.AppSettings["CredentialsInfo"]);
                String FilePath = Path.Combine(FolderPath, "APITokenCredentials");

                credential = GoogleWebAuthorizationBroker.AuthorizeAsync(
                    GoogleClientSecrets.FromStream(stream).Secrets,
                    Scopes,
                    "user",
                    CancellationToken.None,
                    new FileDataStore(FilePath, true)).Result;
            }

            GmailService service = new GmailService(new BaseClientService.Initializer()
            {
                HttpClientInitializer = credential,
                ApplicationName = ApplicationName,
            });

            return service;
        }

        public static void MsgMarkAsread(string hostEmailAddress, string msgId)
        {
            //MESSAGE MARKS AS READ AFTER READING MESSAGE
            ModifyMessageRequest mods = new ModifyMessageRequest();
            mods.AddLabelIds = null;
            mods.RemoveLabelIds = new List<string> { "UNREAD" };
            GetService().Users.Messages.Modify(mods, hostEmailAddress, msgId).Execute();
        }

        public static List<string> GetAttachments(string userId, string messageId, String outputDir)
        {
            List<string> fileName = new List<string>();

            try
            {
                GmailService gService = GetService();
                Message message = gService.Users.Messages.Get(userId, messageId).Execute();
                IList<MessagePart> parts = message.Payload.Parts;

                foreach (MessagePart part in parts)
                {
                    if (!String.IsNullOrEmpty(part.Filename))
                    {
                        string attId = part.Body.AttachmentId;
                        MessagePartBody attachPart = gService.Users.Messages.Attachments.Get(userId, messageId, attId).Execute();

                        byte[] data = Convert.FromBase64String(attachPart.Data); //Base64ToByte(attachPart.Data);
                        File.WriteAllBytes(Path.Combine(outputDir, part.Filename), data);
                        fileName.Add(part.Filename);
                    }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
            }

            return fileName;
        }

        public static string MsgNestedParts(IList<MessagePart> parts)
        {
            string str = string.Empty;
            if (parts.Count() < 0)
            {
                return string.Empty;
            }
            else
            {
                IList<MessagePart> plainTestMail = parts.Where(x => x.MimeType == "text/plain").ToList();
                IList<MessagePart> attachmentMail = parts.Where(x => x.MimeType == "multipart/alternative").ToList();

                if (plainTestMail.Count() > 0)
                {
                    foreach (MessagePart eachPart in plainTestMail)
                    {
                        if (eachPart.Parts == null)
                        {
                            if (eachPart.Body != null && eachPart.Body.Data != null)
                            {
                                str += eachPart.Body.Data;
                            }
                        }
                        else
                        {
                            return MsgNestedParts(eachPart.Parts);
                        }
                    }
                }

                if (attachmentMail.Count() > 0)
                {
                    foreach (MessagePart eachPart in attachmentMail)
                    {
                        if (eachPart.Parts == null)
                        {
                            if (eachPart.Body != null && eachPart.Body.Data != null)
                            {
                                str += eachPart.Body.Data;
                            }
                        }
                        else
                        {
                            return MsgNestedParts(eachPart.Parts);
                        }
                    }
                }

                return str;
            }
        }

        public static string Base64Decode(string base64Test)
        {
            string encodeText = string.Empty;

            //STEP-1: Replace all special character od base64Test
            encodeText = base64Test.Replace("-", "+");
            encodeText = encodeText.Replace("_", "/");
            encodeText = encodeText.Replace(" ", "+");
            encodeText = encodeText.Replace("=", "+");

            //STEP-2: Fixed invalid length of base64Test
            if (encodeText.Length % 4 > 0) { encodeText += new string('=', 4 - encodeText.Length % 4); }
            else if (encodeText.Length % 4 == 0)
            {
                encodeText = encodeText.Substring(0, encodeText.Length - 1);
                if (encodeText.Length % 4 > 0) { encodeText += new string('+', 4 - encodeText.Length % 4); }
            }

            //Encoding to UTF-8 before create byte array
            encodeText = Encoding.UTF8.GetString(Convert.FromBase64String(encodeText));

            //STEP-3: Convert to Byte array
            //byte[] byteArray = Convert.FromBase64String(encodeText);

            //STEP-4: Encoding to UTF-8 format
            //return Encoding.UTF8.GetString(byteArray);

            return encodeText;
        }

        public static void SeparateBodyAndSaveToDB(string body, DateTime emailDate)
        {
            try
            {
                Info i = new Info();

                // Split the input string by line breaks and create a dictionary
                Dictionary<string, string> keyValuePairs = body
                    .Split(new[] { '\n', '\r' }, StringSplitOptions.RemoveEmptyEntries)
                    .Select(line => line.Split(':'))
                    .Where(parts => parts.Length == 2)
                    .ToDictionary(parts => parts[0], parts => parts[1]);

                #region Assigning Variables

                // Print the key-value pairs
                foreach (var kvp in keyValuePairs)
                {
                    //Clean spaces
                    string keyString = kvp.Key.Trim();
                    string valueString = kvp.Value.Trim();

                    switch (keyString)
                    {
                        case "Nome":
                            i.Nome = valueString;
                            break;
                        case "Nome no Crachá":
                            i.NomeCracha = valueString;
                            break;
                        case "Sexo/Gênero":
                            i.SexoGenero = valueString;
                            break;
                        case "Raça/Cor/Etnia":
                            i.RacaCorEtnia = valueString;
                            break;
                        case "Instituição":
                            i.Instituicao = valueString;
                            break;
                        case "Documento":
                            i.Documento = valueString;
                            break;
                        case "Número do Documento Selecionado":
                            i.NumeroDocumento = valueString;
                            break;
                        case "Profissão":
                            i.Profissao = valueString;
                            break;
                        case "Endereço":
                            i.Endereco = valueString;
                            break;
                        case "Bairro":
                            i.Bairro = valueString;
                            break;
                        case "Cep":
                            i.Cep = valueString;
                            break;
                        case "Município":
                            i.Municipio = valueString;
                            break;
                        case "Cidade":
                            i.Cidade = valueString;
                            break;
                        case "País":
                            i.Pais = valueString;
                            break;
                        case "número de celular":
                            i.NumeroCelular = valueString;
                            break;
                        case "Email":
                            i.Email = valueString;
                            break;
                        case "Tem Deficiência":
                            i.TemDeficiencia = valueString;
                            break;
                        case "Descrição da deficiência, caso tenha":
                            i.DescricaoDeficiencia = valueString;
                            break;
                        case "Participará com acompanhante":
                            i.ComAcompanhante = valueString;
                            break;
                        case "Precisa de Recursos de Acessibilidade no Evento":
                            i.PrecisaRecursos = valueString;
                            break;
                        case "Descrição Recurso de Acessbilidade caso precise":
                            i.DescricaoRecurso = valueString;
                            break;
                        case "Categoria":
                            i.Categoria = valueString;
                            break;
                        default:
                            break;
                    }
                    Console.WriteLine($"Key: {keyString}, Value: {valueString}");
                }

                keyValuePairs.Clear();

                #endregion

                //Save To Database
                Guid newId = Guid.NewGuid();
                var connection = new SqlConnection(@"Server=10.5.90.31;Database=CONVISA_MAIL;User ID=sa;Password=!viS@.2022.At!;");

                var sql = "INSERT INTO Inscricoes (Id, Nome, NomeCracha, SexoGenero, RacaCorEtnia, " +
                    "Instituicao, Documento, NumeroDocumento, Profissao, Endereco, Bairro, Cep, Municipio, Cidade, Pais, " +
                    "NumeroCelular, Email, TemDeficiencia, DescricaoDeficiencia, ComAcompanhante, PrecisaRecursos, " +
                    "DescricaoRecurso, Categoria, Data) VALUES ('"+ newId +"','"+ i.Nome +"','"+ i.NomeCracha +"','"+ i.SexoGenero + "','"+
                    i.RacaCorEtnia +"','"+ i.Instituicao +"','"+ i.Documento +"','"+ i.NumeroDocumento + "','"+ i.Profissao +"','"+
                    i.Endereco +"','"+ i.Bairro +"','"+ i.Cep +"','"+ i.Municipio +"','"+ i.Cidade +"','"+ i.Pais +"','"+
                    i.NumeroCelular +"','"+ i.Email +"','"+ i.TemDeficiencia +"','"+ i.DescricaoDeficiencia +"','"+
                    i.ComAcompanhante +"','"+ i.PrecisaRecursos +"','"+ i.DescricaoRecurso +"','"+ i.Categoria +"', '"+ emailDate.ToLocalTime() +"')";

                IEnumerable<Info> results = null;

                if (i.Nome != string.Empty || i.Nome != null)
                {
                    results = connection.Query<Info>(sql);

                    Console.WriteLine("Usuário inserido com sucesso! => i: " + i.Nome + " - " + i.NomeCracha + " - " + i.SexoGenero + " - " + i.RacaCorEtnia + " - " +
                        i.Instituicao + " - " + i.Documento + " - " + i.NumeroDocumento + " - " + i.Profissao + " - " + i.Endereco + " - " + i.Bairro + " - " + i.Cep + " - " + i.Municipio + " - " + i.Cidade + " - " +
                        i.Pais + " - " + i.NumeroCelular + " - " + i.Email + " - " + i.TemDeficiencia + " - " + i.DescricaoDeficiencia + " - " + i.ComAcompanhante + " - " + i.PrecisaRecursos + " - " +
                        i.DescricaoRecurso + " - " + i.Categoria + " - " + emailDate.ToLocalTime());
                }
                else
                {
                    Console.WriteLine("Usuário não inserido! => i: " + newId + "");
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
            }
        }
    }
}
