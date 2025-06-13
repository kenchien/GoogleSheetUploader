using System;
using System.IO;
using System.Threading.Tasks;
using Google.Apis.Auth.OAuth2;
using Google.Apis.Auth.OAuth2.Flows;
using Google.Apis.Auth.OAuth2.Responses;
using Google.Apis.Util.Store;
using Common.Helper;

namespace GoogleSheetUploader.Helper
{
    public class TokenHelper
    {
        private readonly string _secretJson;
        private readonly string _user;
        private readonly string[] _scopes;
        private readonly IDataStore _dataStore;

        public TokenHelper(string secretJson, string user, string[] scopes)
        {
            _secretJson = secretJson ?? throw new ArgumentNullException(nameof(secretJson));
            _user = user ?? throw new ArgumentNullException(nameof(user));
            _scopes = scopes ?? throw new ArgumentNullException(nameof(scopes));
            _dataStore = new FileDataStore("Google.Apis.Auth", true);
        }

        public async Task<UserCredential> GetCredentialAsync()
        {
            try
            {
                var credential = await GoogleWebAuthorizationBroker.AuthorizeAsync(
                    GoogleClientSecrets.FromFile(_secretJson).Secrets,
                    _scopes,
                    _user,
                    System.Threading.CancellationToken.None,
                    _dataStore
                );

                if (credential.Token.IsStale)
                {
                    await credential.RefreshTokenAsync(System.Threading.CancellationToken.None);
                }

                return credential;
            }
            catch (Exception ex)
            {
                LogHelper.Error("取得 Google API 認證時發生錯誤", ex);
                throw;
            }
        }

        public async Task RefreshTokenAsync(UserCredential credential)
        {
            try
            {
                if (credential.Token.IsExpired(credential.Flow.Clock))
                {
                    await credential.RefreshTokenAsync(CancellationToken.None);

                    // 更新儲存的 token
                    var tokenJson = Newtonsoft.Json.JsonConvert.SerializeObject(credential.Token);
                    await File.WriteAllTextAsync(_secretJson, tokenJson);
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"更新憑證時發生錯誤: {ex.Message}");
                throw;
            }
        }
    }
} 