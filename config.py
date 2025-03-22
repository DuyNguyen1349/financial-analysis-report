import os

class Config:
    """Base configuration class"""
    SECRET_KEY = os.environ.get('SECRET_KEY') or 'your-secret-key-here'
    DATA_DIR = os.path.join(os.path.abspath(os.path.dirname(__file__)), 'data')
    DEBUG = False
    TESTING = False
    
    # Application specific settings
    APP_NAME = "Báo Cáo Phân Tích Tài Chính Chuyên Nghiệp"
    APP_VERSION = "1.0.0"
    
    # Financial analysis settings
    DEFAULT_YEARS = [2020, 2021, 2022, 2023, 2024]
    DEFAULT_METRICS = [
        'ROA (%)', 'ROE (%)', 'ROS (%)', 'EBIT Margin (%)', 'EBITDA Margin (%)',
        'Gross Profit Margin (%)', 'Current Ratio', 'Quick Ratio',
        'D/A (%)', 'D/E (%)', 'E/A (%)', 'Interest Coverage Ratio'
    ]
    
    # Chart colors
    CHART_COLORS = [
        'rgba(255, 99, 132, 0.7)',  # Red
        'rgba(54, 162, 235, 0.7)',  # Blue
        'rgba(255, 206, 86, 0.7)',  # Yellow
        'rgba(75, 192, 192, 0.7)',  # Green
        'rgba(153, 102, 255, 0.7)', # Purple
        'rgba(255, 159, 64, 0.7)'   # Orange
    ]

class DevelopmentConfig(Config):
    """Development configuration"""
    DEBUG = True
    
    # Enable more detailed logging
    LOGGING_LEVEL = 'DEBUG'
    
    # Development-specific settings
    TEMPLATES_AUTO_RELOAD = True
    
class TestingConfig(Config):
    """Testing configuration"""
    TESTING = True
    DEBUG = True
    
    # Use a separate data directory for testing
    DATA_DIR = os.path.join(os.path.abspath(os.path.dirname(__file__)), 'test_data')
    
class ProductionConfig(Config):
    """Production configuration"""
    # Production specific settings
    DEBUG = False
    TESTING = False
    
    # Ensure secure operations
    SESSION_COOKIE_SECURE = True
    REMEMBER_COOKIE_SECURE = True
    
    # Higher level of security for CSRF protection
    WTF_CSRF_ENABLED = True
    
    # Logging configuration
    LOGGING_LEVEL = 'ERROR'

# Configuration dictionary
config = {
    'development': DevelopmentConfig,
    'testing': TestingConfig,
    'production': ProductionConfig,
    'default': DevelopmentConfig
}

def get_config(config_name='default'):
    """Get configuration by name"""
    return config.get(config_name, config['default'])