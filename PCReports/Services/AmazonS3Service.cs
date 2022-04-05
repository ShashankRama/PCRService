using Amazon.S3;
using Amazon.S3.Model;

namespace PCReports.Services
{
    public class AmazonS3Service
    {
        public AmazonS3Service(IConfiguration configuration)
        {
            Configuration = configuration;
            accessKey = Configuration["AWSS3Settings:DE_AWSAccessKey_PhysicsCore"];
            secretKey = Configuration["AWSS3Settings:DE_AWSSecretKey_PhysicsCore"];
            serviceUrl = Configuration["AWSS3Settings:DES3ServiceUrl"];
            bucketName = Configuration["AWSS3Settings:DEAWSBucket_PhysicsCore"];
            bucketServer = Configuration["AWSS3Settings:DE_AWSBucket_PhysicsCore_Server"];
        }

        public async Task<bool> uploadFileToS3Async(string localFilePath, string fileURL)
        {
            try
            {
                AmazonS3Client s3Client = new AmazonS3Client(
                accessKey,
                secretKey,
                new AmazonS3Config
                {
                    ServiceURL = serviceUrl
                });

                string key = @"reports/" + bucketServer + @"/" + fileURL;

                var putRequest = new PutObjectRequest
                {
                    BucketName = bucketName,
                    Key = key,
                    FilePath = localFilePath,
                    ContentType = "application/pdf"
                };

                await s3Client.PutObjectAsync(putRequest);

                return true; //indicate that the file was sent
            }
            catch (AmazonS3Exception e)
            {
                throw e;
            }
        }

        private string accessKey;
        private string secretKey;
        private string serviceUrl;
        private string bucketName;
        private string bucketServer;
        private readonly IConfiguration Configuration;
    }
}