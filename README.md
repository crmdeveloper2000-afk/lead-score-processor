# Lead Score Processing Service

A Flask web service that processes lead data and generates PowerPoint presentations with charts and analytics.

## Features

- Accepts lead data via HTTP POST requests
- Downloads PowerPoint templates from Zoho WorkDrive
- Generates charts and analytics based on lead scores
- Uploads processed presentations back to Zoho WorkDrive
- Attaches files to lead records in Zoho CRM

## API Endpoints

### Health Check
```
GET /
```
Returns service status and health information.

### Process Lead
```
POST /process-lead
Content-Type: application/json
```

Processes lead data and generates a PowerPoint presentation.

**Request Body Example:**
```json
{
  "Email": "test@example.com",
  "Organization": "Test Organization",
  "First_Name": "John",
  "Last_Name": "Doe",
  "Lead_ID": "123456789",
  "Domain_1_Sum": "6",
  "Domain_2_Sum": "5",
  "Domain_3_Sum": "2",
  "Domain_4_Sum": "8",
  "Total_Sum": "21",
  "Governance_Q1": "4. Breed gedragen visie vertaald naar programma's",
  "Governance_Q2": "2. Incidentele betrokkenheid bij projecten",
  "Structuur_Q1": "2. Incidentieel overleg of project",
  "Structuur_Q2": "3. Een regionaal dashboard of platform in ontwikkeling",
  "Proces_Q1": "1. Proces gebaseerd op aanbod en interne structuur",
  "Proces_Q2": "1. Geen gestructureerde verbetering",
  "Uitkomsten_en_sturing_Q1": "3. Per populatie of programma wordt op uitkomst KPI's gestuurd",
  "Uitkomsten_en_sturing_Q2": "5. Data gekoppeld aan leren, verbeteren en verantwoorden",
  "Created_Date": "2025/08/28"
}
```

**Response Example:**
```json
{
  "success": true,
  "message": "PPT processed and uploaded successfully and attached to Lead record",
  "lead_id": "123456789",
  "upload_result": {
    "success": true,
    "file_id": "file_id_here",
    "download_url": "https://download.url.here",
    "filename": "processed_output.pptx"
  },
  "attachment_result": {
    "success": true,
    "attachment_id": "attachment_id_here"
  }
}
```

## Deployment to Render.com

1. **Create a new Web Service on Render.com**
   - Connect your GitHub repository
   - Select "Web Service"
   - Choose your repository

2. **Configure Environment Variables**
   Set the following environment variables in your Render.com dashboard:
   - `ZOHO_REFRESH_TOKEN`: Your Zoho refresh token
   - `ZOHO_CLIENT_ID`: Your Zoho client ID
   - `ZOHO_CLIENT_SECRET`: Your Zoho client secret
   - `TEMPLATE_DOWNLOAD_URL`: URL to download the PowerPoint template

3. **Deploy**
   - Render.com will automatically detect the `requirements.txt` and deploy
   - The service will be available at your assigned Render.com URL

## Local Development

1. **Install Dependencies**
   ```bash
   pip install -r requirements.txt
   ```

2. **Set Environment Variables**
   ```bash
   export ZOHO_REFRESH_TOKEN="your_token_here"
   export ZOHO_CLIENT_ID="your_client_id_here"
   export ZOHO_CLIENT_SECRET="your_client_secret_here"
   export TEMPLATE_DOWNLOAD_URL="your_template_url_here"
   ```

3. **Run the Application**
   ```bash
   python Lead-Score.py
   ```

   The service will be available at `http://localhost:5000`

## Files Structure

- `Lead-Score.py`: Main Flask application
- `requirements.txt`: Python dependencies
- `Procfile`: Process configuration for deployment
- `render.yaml`: Render.com specific configuration
- `runtime.txt`: Python version specification
- `README.md`: This documentation

## Usage Example

```bash
curl -X POST https://your-render-app.onrender.com/process-lead \
  -H "Content-Type: application/json" \
  -d '{
    "Email": "test@example.com",
    "Organization": "Test Org",
    "First_Name": "John",
    "Last_Name": "Doe",
    "Lead_ID": "123456789",
    "Domain_1_Sum": "6",
    "Domain_2_Sum": "5",
    "Domain_3_Sum": "2",
    "Domain_4_Sum": "8"
  }'
```
