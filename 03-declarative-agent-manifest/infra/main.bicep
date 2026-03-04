// ============================================================================
// Pattern 3: Foundry Analysis Endpoint — Azure Infrastructure
// ============================================================================
// Deploys:
//   - App Service Plan (B1)
//   - App Service (Python 3.11) with system-assigned managed identity
//   - App settings pointing to Foundry project
//
// Usage:
//   az deployment group create \
//     --resource-group <rg-name> \
//     --template-file main.bicep \
//     --parameters \
//       foundryProjectEndpoint='https://<acct>.services.ai.azure.com/api/projects/<proj>' \
//       foundryAgentId='asst_xxxxxxxxxxxxxxxxxxxx' \
//       tenantId='<tenant-id>' \
//       clientId='<client-id>'
// ============================================================================

@description('Azure region for all resources')
param location string = resourceGroup().location

@description('Base name for resources (used as prefix)')
param baseName string = 'foundry-endpoint'

@description('Foundry project endpoint URL')
param foundryProjectEndpoint string

@description('Foundry agent ID (starts with asst_)')
param foundryAgentId string

@description('Entra tenant ID for token validation')
param tenantId string

@description('Entra app registration client ID for token validation')
param clientId string

@description('App Service Plan SKU')
param skuName string = 'B1'

@description('Resource tags')
param tags object = {
  project: 'foundry-sharepoint-obo-sample'
  pattern: 'pattern-3-declarative-agent'
}

// ---------------------------------------------------------------------------
// App Service Plan
// ---------------------------------------------------------------------------
resource appServicePlan 'Microsoft.Web/serverfarms@2023-12-01' = {
  name: '${baseName}-plan'
  location: location
  tags: tags
  sku: {
    name: skuName
  }
  kind: 'linux'
  properties: {
    reserved: true // Required for Linux
  }
}

// ---------------------------------------------------------------------------
// App Service
// ---------------------------------------------------------------------------
resource appService 'Microsoft.Web/sites@2023-12-01' = {
  name: baseName
  location: location
  tags: tags
  identity: {
    type: 'SystemAssigned'
  }
  properties: {
    serverFarmId: appServicePlan.id
    httpsOnly: true
    siteConfig: {
      linuxFxVersion: 'PYTHON|3.11'
      alwaysOn: true
      minTlsVersion: '1.2'
      ftpsState: 'Disabled'
      healthCheckPath: '/health'
      appSettings: [
        {
          name: 'FOUNDRY_PROJECT_ENDPOINT'
          value: foundryProjectEndpoint
        }
        {
          name: 'FOUNDRY_AGENT_ID'
          value: foundryAgentId
        }
        {
          name: 'AZURE_TENANT_ID'
          value: tenantId
        }
        {
          name: 'AZURE_CLIENT_ID'
          value: clientId
        }
        {
          name: 'USE_MANAGED_IDENTITY'
          value: 'true'
        }
        {
          name: 'FOUNDRY_TIMEOUT_SECONDS'
          value: '90'
        }
        {
          name: 'SCM_DO_BUILD_DURING_DEPLOYMENT'
          value: 'true'
        }
      ]
      appCommandLine: 'gunicorn --bind 0.0.0.0:8000 --timeout 120 --workers 2 app:app'
    }
  }
}

// ---------------------------------------------------------------------------
// Outputs
// ---------------------------------------------------------------------------
output appServiceName string = appService.name
output appServiceUrl string = 'https://${appService.properties.defaultHostName}'
output principalId string = appService.identity.principalId
output resourceId string = appService.id

// ---------------------------------------------------------------------------
// NOTE: Foundry RBAC role assignment
// ---------------------------------------------------------------------------
// After deployment, grant the App Service's managed identity the
// "Azure AI Developer" role on your Foundry project resource:
//
//   az role assignment create \
//     --assignee <principalId from output> \
//     --role "Azure AI Developer" \
//     --scope /subscriptions/<sub>/resourceGroups/<rg>/providers/Microsoft.MachineLearningServices/workspaces/<foundry-project>
//
// This cannot be done in this Bicep file because the Foundry project
// resource is typically in a different resource group.
