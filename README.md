# Report Builder for Cisco Intersight

This tool builds Excel reports of data in Cisco Intersight.

## Setup
1. Prepare a host with Python 3.x
2. Download a release of the tool onto your host and unzip it
3. Install the required modules via 'pip install - r requirements.txt'
4. Create a configuration file in the main directory of the downloaded folder, and add other files as required

## Configuration File
Before you can run the this tool, you will have to create a configuration file. The basic template for it looks like this:
```yaml
intersight:
    endpoint: "https://intersight.com"
    api_key: "API KEY"
    secret_key_path: "./SecretKey.txt"
logging:
    level: "INFO"
```

You can use this template to get started with building your own configuration file. Below, all sections will be explained in detail.

### intersight
The *intersight* settings control what Intersight instance will be used, and how you authenticate with it.
| Key             | Default | Description                                  |
|-----------------|---------|----------------------------------------------|
| endpoint        |         | The URL of the Intersight instance (string)  |
| api_key         |         | The API key used for authentication (string) |
| secret_key_path |         | The filepath of your SecretKey (string)      |

Keep in mind that the 'secret_key_path' requires you to also add the text file with the SecretKey. You can create the API Key and the SecretKey from the Intersight system settings.

### logging
The *logging* setting determines the log level. Depending on your use case, you can use more or less restrictive rules.
| Key   | Default | Description                                                         |
|-------|---------|---------------------------------------------------------------------|
| level | INFO    | The log level that should be used for all logs in the tool (string) |

