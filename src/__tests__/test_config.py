"""
HWP 설정 관리 테스트
config.py의 모든 클래스와 함수에 대한 단위 테스트를 포함합니다.
"""

import pytest
from unittest.mock import Mock, MagicMock, patch, mock_open
import sys
import os
import json
import tempfile
import shutil

# 테스트를 위한 경로 설정
sys.path.append(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from tools.config import (
    HwpConfig,
    ConfigManager,
    get_config,
    update_config,
    save_config,
    reset_config
)
from tools.constants import *


class TestHwpConfig:
    """HwpConfig 클래스 테스트"""
    
    def test_default_initialization(self):
        """기본값으로 초기화"""
        # When
        config = HwpConfig()
        
        # Then
        assert config.auto_connect is True
        assert config.visible is True
        assert config.register_security_module is True
        assert config.default_font == "맑은 고딕"
        assert config.default_font_size == 10
        assert config.default_paper_size == "A4"
        assert config.debug_mode is False
        assert isinstance(config.default_margins, dict)
        assert isinstance(config.allowed_image_formats, list)
    
    def test_custom_initialization(self):
        """사용자 정의 값으로 초기화"""
        # When
        config = HwpConfig(
            auto_connect=False,
            default_font="바탕",
            debug_mode=True,
            batch_size=100
        )
        
        # Then
        assert config.auto_connect is False
        assert config.default_font == "바탕"
        assert config.debug_mode is True
        assert config.batch_size == 100
        # 다른 값들은 기본값 유지
        assert config.visible is True
    
    def test_to_dict(self):
        """설정을 딕셔너리로 변환"""
        # Given
        config = HwpConfig(
            auto_connect=False,
            default_font="굴림",
            debug_mode=True
        )
        
        # When
        result = config.to_dict()
        
        # Then
        assert isinstance(result, dict)
        assert result['auto_connect'] is False
        assert result['default_font'] == "굴림"
        assert result['debug_mode'] is True
        assert 'visible' in result
        assert 'default_margins' in result
    
    def test_from_dict(self):
        """딕셔너리에서 설정 객체 생성"""
        # Given
        data = {
            'auto_connect': False,
            'default_font': '돋움',
            'default_font_size': 12,
            'debug_mode': True,
            'batch_size': 50
        }
        
        # When
        config = HwpConfig.from_dict(data)
        
        # Then
        assert config.auto_connect is False
        assert config.default_font == "돋움"
        assert config.default_font_size == 12
        assert config.debug_mode is True
        assert config.batch_size == 50
    
    def test_save_to_file(self):
        """설정을 파일로 저장"""
        # Given
        config = HwpConfig(default_font="나눔고딕", debug_mode=True)
        
        with tempfile.NamedTemporaryFile(mode='w', delete=False, suffix='.json') as f:
            temp_path = f.name
        
        try:
            # When
            config.save(temp_path)
            
            # Then
            assert os.path.exists(temp_path)
            with open(temp_path, 'r', encoding='utf-8') as f:
                data = json.load(f)
            assert data['default_font'] == "나눔고딕"
            assert data['debug_mode'] is True
        finally:
            os.unlink(temp_path)
    
    def test_load_from_file(self):
        """파일에서 설정 로드"""
        # Given
        data = {
            'auto_connect': False,
            'default_font': '궁서',
            'debug_mode': True,
            'log_level': 'DEBUG'
        }
        
        with tempfile.NamedTemporaryFile(mode='w', delete=False, suffix='.json') as f:
            json.dump(data, f)
            temp_path = f.name
        
        try:
            # When
            config = HwpConfig.load(temp_path)
            
            # Then
            assert config.auto_connect is False
            assert config.default_font == "궁서"
            assert config.debug_mode is True
            assert config.log_level == "DEBUG"
        finally:
            os.unlink(temp_path)
    
    def test_load_from_nonexistent_file(self):
        """존재하지 않는 파일에서 로드 시 기본값 반환"""
        # When
        config = HwpConfig.load("/nonexistent/path/config.json")
        
        # Then
        assert config.auto_connect is True  # 기본값
        assert config.default_font == "맑은 고딕"  # 기본값


class TestConfigManager:
    """ConfigManager 클래스 테스트"""
    
    def setup_method(self):
        """각 테스트 전 ConfigManager 싱글톤 초기화"""
        ConfigManager._instance = None
        ConfigManager._config = None
    
    def test_singleton_pattern(self):
        """싱글톤 패턴 테스트"""
        # When
        manager1 = ConfigManager()
        manager2 = ConfigManager()
        
        # Then
        assert manager1 is manager2
    
    def test_default_config_load(self):
        """기본 설정 로드"""
        # When
        manager = ConfigManager()
        
        # Then
        assert manager.config is not None
        assert isinstance(manager.config, HwpConfig)
        assert manager.config.auto_connect is True
    
    @patch.dict(os.environ, {'HWP_MCP_CONFIG': '/env/config.json'})
    @patch('os.path.exists')
    @patch('tools.config.HwpConfig.load')
    def test_load_config_from_env(self, mock_load, mock_exists):
        """환경 변수에서 설정 파일 경로 확인"""
        # Given
        mock_exists.side_effect = lambda path: path == '/env/config.json'
        mock_config = HwpConfig(debug_mode=True)
        mock_load.return_value = mock_config
        
        # When
        manager = ConfigManager()
        
        # Then
        mock_exists.assert_any_call('/env/config.json')
        mock_load.assert_called_with('/env/config.json')
        assert manager.config.debug_mode is True
    
    @patch('os.path.exists')
    def test_load_config_search_paths(self, mock_exists):
        """여러 경로에서 설정 파일 검색"""
        # Given
        mock_exists.return_value = False
        
        # When
        manager = ConfigManager()
        
        # Then
        # 설정 파일 검색 경로 확인
        assert mock_exists.call_count >= 3
        # 기본 설정 반환
        assert manager.config.auto_connect is True
    
    def test_update_config(self):
        """설정 업데이트"""
        # Given
        manager = ConfigManager()
        
        # When
        manager.update(
            debug_mode=True,
            default_font="바탕",
            batch_size=200
        )
        
        # Then
        assert manager.config.debug_mode is True
        assert manager.config.default_font == "바탕"
        assert manager.config.batch_size == 200
    
    def test_update_invalid_attribute(self):
        """존재하지 않는 속성 업데이트 시도"""
        # Given
        manager = ConfigManager()
        
        # When
        manager.update(invalid_attribute="value")
        
        # Then
        # 예외가 발생하지 않고 무시됨
        assert not hasattr(manager.config, 'invalid_attribute')
    
    def test_save_config(self):
        """설정 저장"""
        # Given
        manager = ConfigManager()
        manager.update(debug_mode=True)
        
        with tempfile.NamedTemporaryFile(mode='w', delete=False, suffix='.json') as f:
            temp_path = f.name
        
        try:
            # When
            manager.save(temp_path)
            
            # Then
            assert os.path.exists(temp_path)
            with open(temp_path, 'r', encoding='utf-8') as f:
                data = json.load(f)
            assert data['debug_mode'] is True
        finally:
            os.unlink(temp_path)
    
    def test_save_config_default_path(self):
        """기본 경로에 설정 저장"""
        # Given
        manager = ConfigManager()
        
        # When
        with patch.object(manager._config, 'save') as mock_save:
            manager.save()
        
        # Then
        mock_save.assert_called_once_with("./hwp_config.json")
    
    def test_reset_config(self):
        """설정 초기화"""
        # Given
        manager = ConfigManager()
        manager.update(debug_mode=True, default_font="굴림")
        
        # When
        manager.reset()
        
        # Then
        assert manager.config.debug_mode is False  # 기본값으로 복원
        assert manager.config.default_font == "맑은 고딕"  # 기본값으로 복원


class TestConvenienceFunctions:
    """편의 함수 테스트"""
    
    def setup_method(self):
        """각 테스트 전 ConfigManager 싱글톤 초기화"""
        ConfigManager._instance = None
        ConfigManager._config = None
    
    def test_get_config(self):
        """get_config 함수 테스트"""
        # When
        config = get_config()
        
        # Then
        assert isinstance(config, HwpConfig)
        assert config is ConfigManager().config
    
    def test_update_config_function(self):
        """update_config 함수 테스트"""
        # When
        update_config(debug_mode=True, log_level="DEBUG")
        
        # Then
        config = get_config()
        assert config.debug_mode is True
        assert config.log_level == "DEBUG"
    
    @patch.object(ConfigManager, 'save')
    def test_save_config_function(self, mock_save):
        """save_config 함수 테스트"""
        # When
        save_config("/path/to/config.json")
        
        # Then
        mock_save.assert_called_once_with("/path/to/config.json")
    
    @patch.object(ConfigManager, 'save')
    def test_save_config_function_default(self, mock_save):
        """save_config 함수 테스트 (기본 경로)"""
        # When
        save_config()
        
        # Then
        mock_save.assert_called_once_with(None)
    
    def test_reset_config_function(self):
        """reset_config 함수 테스트"""
        # Given
        update_config(debug_mode=True, default_font="돋움")
        
        # When
        reset_config()
        
        # Then
        config = get_config()
        assert config.debug_mode is False
        assert config.default_font == "맑은 고딕"


class TestConfigIntegration:
    """설정 관련 통합 테스트"""
    
    def setup_method(self):
        """각 테스트 전 초기화"""
        ConfigManager._instance = None
        ConfigManager._config = None
        self.temp_dir = tempfile.mkdtemp()
    
    def teardown_method(self):
        """각 테스트 후 정리"""
        shutil.rmtree(self.temp_dir, ignore_errors=True)
    
    def test_full_config_lifecycle(self):
        """전체 설정 수명주기 테스트"""
        config_path = os.path.join(self.temp_dir, "test_config.json")
        
        # 1. 기본 설정으로 시작
        config1 = get_config()
        assert config1.debug_mode is False
        
        # 2. 설정 수정
        update_config(
            debug_mode=True,
            default_font="나눔고딕",
            batch_size=100,
            log_level="DEBUG"
        )
        
        # 3. 설정 저장
        save_config(config_path)
        
        # 4. ConfigManager 재시작
        ConfigManager._instance = None
        ConfigManager._config = None
        
        # 5. 환경 변수로 설정 파일 지정
        with patch.dict(os.environ, {'HWP_MCP_CONFIG': config_path}):
            config2 = get_config()
            
            # 저장된 설정이 로드되었는지 확인
            assert config2.debug_mode is True
            assert config2.default_font == "나눔고딕"
            assert config2.batch_size == 100
            assert config2.log_level == "DEBUG"
    
    def test_config_error_handling(self):
        """설정 파일 오류 처리"""
        # 잘못된 JSON 파일 생성
        bad_config_path = os.path.join(self.temp_dir, "bad_config.json")
        with open(bad_config_path, 'w') as f:
            f.write("{ invalid json }")
        
        # 환경 변수로 잘못된 설정 파일 지정
        with patch.dict(os.environ, {'HWP_MCP_CONFIG': bad_config_path}):
            with patch('builtins.print') as mock_print:
                # ConfigManager 재시작
                ConfigManager._instance = None
                ConfigManager._config = None
                
                # 오류가 발생해도 기본 설정이 로드되어야 함
                config = get_config()
                assert config.auto_connect is True  # 기본값
                
                # 오류 메시지 출력 확인
                mock_print.assert_called()