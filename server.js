const express = require('express');
const axios = require('axios');
const xml2js = require('xml2js');
const soap = require('soap');
const cors = require('cors');
const { Client } = require('@microsoft/microsoft-graph-client');
const { ClientSecretCredential } = require('@azure/identity');

// load .env
require('dotenv').config();

const app = express();
const port = process.env.PORT;
const prefix = '/api/v1';

// enable JSON format
app.use(express.json());

// enable CORS
app.use(cors());

/**
 * Fetch civil information from ROP 
 */
app.post(prefix + '/rop/fetch-civil-info', async (req, res) => {
  try {
    const data = req.body;
    const wsdlUrl = 'http://10.14.7.77/ROP-PRO-V-2-DT/rop_service.asmx?wsdl';

    if (data.civilId === undefined || !data.civilId) {
      return res.status(400).json({ error: 'Invalid Civil Id' });
    }

    if (data.expiryDate === undefined || !data.expiryDate) {
      return res.status(400).json({ error: 'Invalid Expiry Date' });
    }

    const args = {
      ServicePassword: process.env.ROP_SERVICE_PASSWORD,
      CivilID: data.civilId,
      ExpiryDate: data.expiryDate
    };

    soap.createClient(wsdlUrl, (err, client) => {
      if (err) {
        return res.status(500).json({ error: 'Error : ' + err });
      }

      // Call the method as it appears in the WSDL
      client.GetByCivilIDFromROP(args, (err, result) => {
        if (err) {
          return res.status(500).json({ error: 'Error : ' + err });
        }

        return res.send(result.GetByCivilIDFromROPResult);
      });
    });
  } catch (error) {
    return res.status(500).json({ error: 'Error : ' + error });
  }
});

/**
 * Send SMS message using TamimahSms
 */
app.post(prefix + '/send-sms', async (req, res) => {
  try {
    const data = req.body;
    const wsdlUrl = 'https://tamimahsms.com/user/BulkPush.asmx?wsdl';

    if (data.message === undefined || !data.message) {
      return res.status(400).json({ error: 'Invalid Message' });
    }

    if (data.mobiles === undefined || !data.mobiles) {
      return res.status(400).json({ error: 'Invalid Mobile Numbers' });
    }

    const args = {
      UserName: process.env.SMS_USERNAME,
      Password: process.env.SMS_PASSWORD,
      Message: data.message,
      Priority: '2',
      Sender: 'CPA',
      SourceRef: 'value2',
      MSISDNs: data.mobiles,
    };

    soap.createClient(wsdlUrl, (error, client) => {
      if (error) {
        return res.status(500).json({ error: 'Error : ' + error });
      }

      client.SendSMS(args, (error, result) => {
        if (error) {
          return res.status(500).json({ error: 'Error : ' + error });
        }

        return res.json(result);
      });
    });
  } catch (error) {
    return res.status(500).json({ error: 'Error : ' + error });
  }
});

/**
 * Send email using Microsoft Azure Graph
 */
app.post(prefix + '/send-email', async (req, res) => {

  try {
    const data = req.body;
    const sender = 'no.reply@cpa.gov.om';

    if (data.recipients === undefined || !data.recipients) {
      return res.status(400).json({ error: 'Invalid recipients email addresses' });
    }

    if (data.subject === undefined || !data.subject) {
      return res.status(400).json({ error: 'Invalid email subject' });
    }

    if (data.content === undefined || !data.content) {
      return res.status(400).json({ error: 'Invalid email content' });
    }

    const message = {
      subject: data.subject,
      body: {
        contentType: 'HTML',
        content: data.content,
      },
      toRecipients: data.recipients.map(email => ({
        emailAddress: {
          address: email,
        },
      })),
      ccRecipients: data.cc.map(email => ({
        emailAddress: {
          address: email,
        },
      })),
      bccRecipients: data.bcc.map(email => ({
        emailAddress: {
          address: email,
        },
      })),
      attachments: data.attachments.map(file => ({
        '@odata.type': '#microsoft.graph.fileAttachment',
        name: file.name,
        contentBytes: file.contentBytes,
      })),
    };

    // Get Graph client and send email
    const tenantId = process.env.EMAIL_TENANT_ID; 
    const clientId = process.env.EMAIL_CLIENT_ID; 
    const clientSecret = process.env.EMAIL_CLIENT_SECRET;

    const credential = new ClientSecretCredential(
      tenantId, 
      clientId, 
      clientSecret 
    );

    const client = Client.initWithMiddleware({
      authProvider: {
        getAccessToken: async () => {
          const tokenResponse = await credential.getToken("https://graph.microsoft.com/.default");
          return tokenResponse.token;
        },
      },
    });
    
    await client.api('/users/' + sender + '/sendMail').post({
      message: message,
    });

    return res.json({ status: 'Email was sent Successfully !' });
  } catch (error) {
    return res.status(500).json({ error: 'Error : ' + error });
  }
});

/**
 * MOCI : login
 */
app.post(prefix + '/moci/login', async (req, res) => {
  try {
    const serviceUrl = 'https://maidan.cpa.gov.om/cpa-web-api-pro/InspectionMobile/Login';
    
    const response = await axios.post(serviceUrl, req.body, {
      headers: {
        'Content-Type': 'application/json'
      }
    });

    return res.send(response.data);
  } catch (error) {
    return res.status(500).json({ error: 'Error : ' + error });
  }
});

/**
 * MOCI : search company
 */
app.get(prefix + '/moci/search-company/:cr_number', async (req, res) => {
  try {
    const serviceUrl = 'https://maidan.cpa.gov.om/cpa-web-api-pro/InspectionMobile/SearchCompany?CRNumber=' + req.params.cr_number;
    const authHeader = req.headers['authorization'];
    
    if (!authHeader || !authHeader.startsWith('Bearer ')) {
      return res.status(401).json({ error: 'Missing or invalid Authorization header' });
    }

    const bearerToken = authHeader.split(' ')[1]; 

    const response = await axios.get(serviceUrl, {
      headers: {
        'Content-Type': 'application/json',
        'Authorization': `Bearer ${bearerToken}`
      }
    });

    return res.send(response.data);
  } catch (error) {
    return res.status(500).json({ error: 'Error : ' + error });
  }
});

/**
 * MOCI : get company data
 */
app.get(prefix + '/moci/get-company-data/:cr_number', async (req, res) => {
  try {
    const serviceUrl = 'https://maidan.cpa.gov.om/cpa-web-api-pro/InspectionMobile/getCompanyData?CRNumber=' + req.params.cr_number;
    const authHeader = req.headers['authorization'];
    
    if (!authHeader || !authHeader.startsWith('Bearer ')) {
      return res.status(401).json({ error: 'Missing or invalid Authorization header' });
    }

    const bearerToken = authHeader.split(' ')[1]; 

    const response = await axios.get(serviceUrl, {
      headers: {
        'Content-Type': 'application/json',
        'Authorization': `Bearer ${bearerToken}`
      }
    });

    return res.send(response.data);
  } catch (error) {
    return res.status(500).json({ error: 'Error : ' + error });
  }
});

/**
 * MOCI : get declared activities
 */
app.get(prefix + '/moci/get-declared-activities/:cr_number', async (req, res) => {
  try {
    const serviceUrl = 'https://maidan.cpa.gov.om/cpa-web-api-pro/InspectionMobile/getDeclaredActivities?CRNumber=' + req.params.cr_number;
    const authHeader = req.headers['authorization'];
    
    if (!authHeader || !authHeader.startsWith('Bearer ')) {
      return res.status(401).json({ error: 'Missing or invalid Authorization header' });
    }

    const bearerToken = authHeader.split(' ')[1]; 

    const response = await axios.get(serviceUrl, {
      headers: {
        'Content-Type': 'application/json',
        'Authorization': `Bearer ${bearerToken}`
      }
    });

    return res.send(response.data);
  } catch (error) {
    return res.status(500).json({ error: 'Error : ' + error });
  }
});

/**
 * MOCI : get places of activities
 */
app.get(prefix + '/moci/get-places-of-activities/:cr_number', async (req, res) => {
  try {
    const serviceUrl = 'https://maidan.cpa.gov.om/cpa-web-api-pro/InspectionMobile/getPlacesOfActivities?CRNumber=' + req.params.cr_number;
    const authHeader = req.headers['authorization'];
    
    if (!authHeader || !authHeader.startsWith('Bearer ')) {
      return res.status(401).json({ error: 'Missing or invalid Authorization header' });
    }

    const bearerToken = authHeader.split(' ')[1]; 

    const response = await axios.get(serviceUrl, {
      headers: {
        'Content-Type': 'application/json',
        'Authorization': `Bearer ${bearerToken}`
      }
    });

    return res.send(response.data);
  } catch (error) {
    return res.status(500).json({ error: 'Error : ' + error });
  }
});

// start the server
app.listen(port, () => {
  console.log(`Server is running on port : ${port}`);
});
