# Weather App

A simple yet powerful Python application that provides real-time weather information for any city worldwide. The application fetches data from [WeatherAPI](https://www.weatherapi.com/) and utilizes the Windows Speech API (SAPI) to read out weather details audibly, making it accessible and user-friendly.

[![License: MIT](https://img.shields.io/badge/License-MIT-yellow.svg)](https://opensource.org/licenses/MIT)
[![Python Version](https://img.shields.io/badge/python-3.6%2B-blue)](https://www.python.org/downloads/)
[![Platform](https://img.shields.io/badge/platform-Windows-lightgrey)](https://www.microsoft.com/windows)

## Table of Contents

- [Features](#features)
- [Prerequisites](#prerequisites)
- [Installation](#installation)
- [Configuration](#configuration)
- [Usage](#usage)
- [API Information](#api-information)
- [How It Works](#how-it-works)
- [Troubleshooting](#troubleshooting)
- [Contributing](#contributing)
- [License](#license)
- [Author](#author)
- [Acknowledgments](#acknowledgments)

## Features

- **Real-time Weather Data**: Fetches current weather information from WeatherAPI
- **Dual Temperature Units**: Displays temperature in both Celsius and Fahrenheit
- **Comprehensive Wind Information**: 
  - Wind speed in kilometers per hour (kph) and miles per hour (mph)
  - Wind direction (N, S, E, W, NE, SE, SW, NW)
- **"Feels Like" Temperature**: Provides the apparent temperature considering wind chill and humidity
- **Text-to-Speech Integration**: Reads out weather information using Windows Speech API (SAPI)
- **Simple CLI Interface**: Easy-to-use command-line interface
- **Environment Variable Support**: Secure API key management using environment variables

## Prerequisites

Before running this application, ensure you have the following installed:

- **Python 3.6 or higher** - [Download Python](https://www.python.org/downloads/)
- **Windows Operating System** - Required for SAPI text-to-speech functionality
- **Internet Connection** - Required to fetch weather data from the API

## Installation

1. **Clone the repository**:
   ```bash
   git clone https://github.com/MuhammadFurqanMohsin25Apr/Weather_app.git
   cd Weather_app
   ```

2. **Install required dependencies**:
   ```bash
   pip install requests pywin32
   ```

   The application requires the following Python packages:
   - `requests` - For making HTTP requests to the WeatherAPI
   - `pywin32` - For Windows Speech API integration

## Configuration

### Setting up WeatherAPI Key

1. **Get your free API key**:
   - Visit [WeatherAPI](https://www.weatherapi.com/)
   - Sign up for a free account
   - Navigate to your dashboard to obtain your API key

2. **Set up environment variable**:

   **Option 1: Temporary (Current Session Only)**
   ```powershell
   # PowerShell
   $env:my_api = "your_api_key_here"
   ```
   ```cmd
   # Command Prompt
   set my_api=your_api_key_here
   ```

   **Option 2: Permanent (Recommended)**
   - Open **System Properties** → **Advanced** → **Environment Variables**
   - Under **User variables**, click **New**
   - Variable name: `my_api`
   - Variable value: Your WeatherAPI key
   - Click **OK** and restart your terminal

   **Option 3: Using Python**
   ```python
   import os
   os.environ['my_api'] = 'your_api_key_here'
   ```

## Usage

1. **Run the application**:
   ```bash
   python app.py
   ```

2. **Enter a city name when prompted**:
   ```
   Enter the name of the city: London
   ```

3. **The application will**:
   - Fetch current weather data for the specified city
   - Display the weather information in the console
   - Read out the weather details using text-to-speech

### Example Output

```
The current temperature of London is 15 degrees in celcius. and in fahrenheit it is 59 degrees. 
the wind speed is 20 kilometers per hour. and in miles per hour it is 12 miles per hour. 
the wind direction is SW. the feels like temperature is 13 degrees in celcius and 55 degrees in fahrenheit.
```

## API Information

This application uses the [WeatherAPI](https://www.weatherapi.com/) service:

- **Endpoint**: `http://api.weatherapi.com/v1/current.json`
- **Free Tier**: 1,000,000 calls/month
- **Data Provided**: Current weather conditions including temperature, wind, humidity, and more
- **Coverage**: Global weather data

### API Response Structure

The application extracts the following data from the API response:
```json
{
  "current": {
    "temp_c": 15.0,
    "temp_f": 59.0,
    "wind_kph": 20.0,
    "wind_mph": 12.4,
    "wind_dir": "SW",
    "feelslike_c": 13.0,
    "feelslike_f": 55.4
  }
}
```

## How It Works

1. **User Input**: The application prompts the user to enter a city name
2. **API Key Retrieval**: Retrieves the API key from the environment variable `my_api`
3. **API Request**: Sends a GET request to WeatherAPI with the city name and API key
4. **Data Parsing**: Parses the JSON response to extract weather information
5. **Speech Synthesis**: Constructs a natural language sentence with weather details
6. **Audio Output**: Uses Windows SAPI to read out the weather information

## Troubleshooting

### Common Issues

**Issue**: `KeyError: 'current'`
- **Cause**: Invalid city name or API request failure
- **Solution**: Verify the city name spelling and check your internet connection

**Issue**: `NameError: name 'my_api' is not defined`
- **Cause**: Environment variable not set
- **Solution**: Ensure you've set the `my_api` environment variable correctly

**Issue**: `ModuleNotFoundError: No module named 'win32com'`
- **Cause**: pywin32 package not installed
- **Solution**: Run `pip install pywin32`

**Issue**: Text-to-speech not working
- **Cause**: Windows Speech API not available
- **Solution**: Ensure you're running on Windows and SAPI is properly configured

**Issue**: API rate limit exceeded
- **Cause**: Too many requests to WeatherAPI
- **Solution**: Wait for the rate limit to reset or upgrade your API plan

## Contributing

Contributions are welcome! Here's how you can help:

1. **Fork the repository**
2. **Create a feature branch**:
   ```bash
   git checkout -b feature/AmazingFeature
   ```
3. **Commit your changes**:
   ```bash
   git commit -m 'Add some AmazingFeature'
   ```
4. **Push to the branch**:
   ```bash
   git push origin feature/AmazingFeature
   ```
5. **Open a Pull Request**

### Ideas for Contribution

- Add support for weather forecasts (5-day, 7-day)
- Implement a graphical user interface (GUI)
- Add support for other operating systems (Linux, macOS)
- Include additional weather metrics (humidity, pressure, UV index)
- Add unit tests
- Create a config file for customization
- Support multiple cities in a single run

## License

This project is licensed under the MIT License - see the [LICENSE](LICENSE) file for details.

## Author

**Muhammad Furqan Mohsin**

- GitHub: [@MuhammadFurqanMohsin25Apr](https://github.com/MuhammadFurqanMohsin25Apr)

## Acknowledgments

- [WeatherAPI](https://www.weatherapi.com/) for providing the weather data service
- Microsoft for the Windows Speech API (SAPI)
- The Python community for excellent libraries and tools

---

**Note**: This application requires Windows OS for the text-to-speech functionality. If you're using Linux or macOS, you can modify the code to use alternative TTS libraries such as `pyttsx3` or `gTTS`.
