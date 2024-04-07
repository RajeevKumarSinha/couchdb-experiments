// Importing necessary libraries
const AWS = require('aws-sdk');
const fs = require('fs');
const XLSX = require('xlsx');
const axios = require('axios');
require('dotenv').config();

// Creating a new S3 instance
const s3 = new AWS.S3({
  accessKeyId: process.env.AWS_ACCESS_KEY_ID,
  secretAccessKey: process.env.AWS_SECRET_ACCESS_KEY,
  region: process.env.AWS_REGION
});

// Function to upload a file to S3
const uploadToS3 = async (glbLink, bucketName, glbName) => {
  // Fetching the file from the provided link
  const response = await axios.get(glbLink, { responseType: 'arraybuffer' });
    // Extracting the file name from the link
    const fileName = glbLink.split('/').pop();
  // Setting up the parameters for the S3 upload
  const params = {
    Bucket: bucketName,
    Key: `glb_files/${fileName}`,
    Body: response.data,
    ContentType: 'model/gltf-binary',
    ACL: 'public-read' // Make the uploaded file publically accessible
  };
  // Debugging the S3_BUCKET_NAME value
  console.log('S3_BUCKET_NAME:', process.env.S3_BUCKET_NAME);
  // Uploading the file to S3
  const data = await s3.upload(params).promise();
  // Returning the S3 link of the uploaded file
  return data.Location;
};

// Function to process the Excel file
const processExcelFile = async () => {
  // Reading the Excel file
  const workbook = XLSX.readFile('output.xlsx');
  const sheet_name_list = workbook.SheetNames;
  const data = XLSX.utils.sheet_to_json(workbook.Sheets[sheet_name_list[0]]);

  // Iterating over each row in the Excel file
  for (let i = 0; i < data.length; i++) {
    // Check if the current row already contains an S3 link
    if (!data[i]['s3 link']) {
      // Fetching the glb link from the current row
      const glbLink = data[i]['Glb Link'];
      // Uploading the glb file to S3 and getting the S3 link
      const s3Link = await uploadToS3(glbLink, process.env.S3_BUCKET_NAME, data[i]['Glb Filename']);
      // Adding the S3 link to the current row
      data[i]['s3 link'] = s3Link;
      // Delaying the process for 3 seconds
      await new Promise(resolve => setTimeout(resolve, 3000));
      console.log(`${i+1} file Uploaded Successfully`)
      // Updating the current row of the output.xlsx file with the s3 link
      XLSX.utils.sheet_to_json(workbook.Sheets[sheet_name_list[0]])[i]['s3 link'] = s3Link;
      // Writing the updated Excel file
      XLSX.writeFile(workbook, 'output.xlsx');
    } else {
      console.log(`S3 link already exists in row ${i+1}, skipping upload`);
    }
  }

};

// Processing the Excel file
processExcelFile();

