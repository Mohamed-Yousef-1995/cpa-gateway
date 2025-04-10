const express = require('express');
const axios = require('axios');
const xml2js = require('xml2js');
const soap = require('soap');
const { Client } = require('@microsoft/microsoft-graph-client');
const { ClientSecretCredential } = require('@azure/identity');

const app = express();
const port = 3000;
const prefix = '/api/v1';

// enable JSON format
app.use(express.json());

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
      ServicePassword: 
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

        // convert XML to JSON
        xml2js.parseString(result, { explicitArray: false, ignoreAttrs: true }, (err, result) => {
          if (err) {
            return res.status(500).json({ error: 'Error : ' + err });
          }

          res.json(result);
        });
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

    if (data.mobile === undefined || !data.mobile) {
      return res.status(400).json({ error: 'Invalid Mobile Number' });
    }

    const args = {
      
      Message: data.message,
      Priority: '2',
      Sender: 'CPA',
      SourceRef: 'value2',
      MSISDNs: data.mobile,
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

    if (data.receiver === undefined || !data.receiver) {
      return res.status(400).json({ error: 'Invalid receiver\'s email address' });
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
        contentType: 'Text',
        content: data.content,
      },
      toRecipients: [
        {
          emailAddress: {
            address: data.receiver,
          },
        },
      ],
    };

    // Get Graph client and send email
    const credential = new ClientSecretCredential(
      
    );

    const client = Client.initWithMiddleware({
      authProvider: credential,
    });

    await client.api('/users/' + sender + '/sendMail').post({
      message: message,
    });
  } catch (error) {
    return res.status(500).json({ error: 'Error : ' + error });
  }
});

// start the server
app.listen(port, () => {
  console.log(`Server is running on port : ${port}`);
});
