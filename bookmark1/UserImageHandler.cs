using System;
using System.Data.SqlClient;
using System.IO;

namespace UserImageHandler
{
    public class UserImageHandler
    {
        public static void InsertUserWithImage(string userName, string filePath, string connectionString)
        {
            try
            {
                // Validate file exists
                if (!File.Exists(filePath))
                {
                    Console.WriteLine("Error: File not found.");
                    return;
                }

                // Convert the image to a hexadecimal string
                string hexImageData = ConvertImageToHex(filePath);

                // Insert data into the database
                using (SqlConnection connection = new SqlConnection(connectionString))
                {
                    string query = "INSERT INTO CentralUserInfo.dbo.users (user_name, FIRST_SIGNATURE) VALUES (@UserName, @FirstSignature)";

                    using (SqlCommand command = new SqlCommand(query, connection))
                    {
                        command.Parameters.AddWithValue("@UserName", userName);
                        command.Parameters.AddWithValue("@FirstSignature", hexImageData);

                        connection.Open();
                        int rowsAffected = command.ExecuteNonQuery();

                        if (rowsAffected > 0)
                        {
                            Console.WriteLine("User and image data inserted successfully.");
                        }
                        else
                        {
                            Console.WriteLine("Error: No rows affected.");
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }

        private static string ConvertImageToHex(string filePath)
        {
            // Read the image file as bytes
            byte[] imageBytes = File.ReadAllBytes(filePath);

            // Convert bytes to hexadecimal string
            return BitConverter.ToString(imageBytes).Replace("-", "");
        }
    }
}
