"""
HWP MCP 프로젝트의 설정 관리 파일
런타임 설정과 환경별 구성을 관리합니다.
"""

import os
import json
from typing import Dict, Any, Optional
from dataclasses import dataclass, asdict, field
from .constants import *

try:
    from .error_handling_guide import validate_file_path
except ImportError:
    # validate_file_path를 위한 간단한 fallback
    def validate_file_path(path: str, must_exist: bool = False) -> str:
        abs_path = os.path.abspath(path)
        if must_exist and not os.path.exists(abs_path):
            raise FileNotFoundError(f"파일을 찾을 수 없습니다: {abs_path}")
        return abs_path

@dataclass
class HwpConfig:
    """HWP MCP 설정 클래스"""
    
    # 일반 설정
    auto_connect: bool = True              # 자동으로 HWP 연결
    visible: bool = True                   # HWP 창 표시 여부
    register_security_module: bool = True  # 보안 모듈 자동 등록
    security_module_path: str = SECURITY_MODULE_DEFAULT_PATH
    
    # 성능 설정
    batch_size: int = BATCH_CHUNK_SIZE     # 배치 작업 크기
    timeout: int = DEFAULT_TIMEOUT         # 기본 타임아웃 (초)
    retry_count: int = MAX_RETRY_COUNT     # 재시도 횟수
    retry_delay: int = RETRY_DELAY         # 재시도 지연 시간 (초)
    
    # 기본값 설정
    default_font: str = "맑은 고딕"        # 기본 글꼴
    default_font_size: int = 10            # 기본 글꼴 크기 (pt)
    default_paper_size: str = "A4"         # 기본 용지 크기
    default_orientation: str = "portrait"  # 기본 용지 방향
    default_margins: Dict[str, int] = field(default_factory=lambda: DEFAULT_MARGINS.copy())
    
    # 표 설정
    table_max_rows: int = TABLE_MAX_ROWS
    table_max_cols: int = TABLE_MAX_COLS
    table_default_style: str = "default"
    
    # 이미지 설정
    allowed_image_formats: list = field(default_factory=lambda: ALLOWED_IMAGE_FORMATS.copy())
    image_embed_mode: int = IMAGE_EMBED_MODE
    image_max_width: int = 0   # 0: 제한 없음
    image_max_height: int = 0  # 0: 제한 없음
    
    # PDF 설정
    pdf_default_quality: str = "high"
    pdf_include_fonts: bool = True
    pdf_include_comments: bool = False
    
    # 로깅 설정
    log_level: str = "INFO"
    log_file: str = "hwp_mcp.log"
    log_max_size: int = 10 * 1024 * 1024  # 10MB
    log_backup_count: int = 5
    
    # 디버그 설정
    debug_mode: bool = False
    show_hwp_errors: bool = True
    save_temp_files: bool = False
    temp_dir: str = "./temp"
    
    def to_dict(self) -> Dict[str, Any]:
        """설정을 딕셔너리로 변환"""
        return asdict(self)
    
    @classmethod
    def from_dict(cls, data: Dict[str, Any]) -> 'HwpConfig':
        """딕셔너리에서 설정 객체 생성"""
        return cls(**data)
    
    def save(self, path: str):
        """설정을 JSON 파일로 저장"""
        # 파일 경로 검증 및 정규화
        validated_path = validate_file_path(path, must_exist=False)
        
        # 디렉토리가 없으면 생성
        dir_path = os.path.dirname(validated_path)
        if dir_path and not os.path.exists(dir_path):
            os.makedirs(dir_path)
        
        with open(validated_path, 'w', encoding='utf-8') as f:
            json.dump(self.to_dict(), f, indent=2, ensure_ascii=False)
    
    @classmethod
    def load(cls, path: str) -> 'HwpConfig':
        """JSON 파일에서 설정 로드"""
        # 파일 경로 검증 및 정규화
        validated_path = validate_file_path(path, must_exist=True)
        
        with open(validated_path, 'r', encoding='utf-8') as f:
            data = json.load(f)
            return cls.from_dict(data)


class ConfigManager:
    """설정 관리자 클래스"""
    
    _instance: Optional['ConfigManager'] = None
    _config: Optional[HwpConfig] = None
    
    def __new__(cls):
        """싱글톤 패턴 구현"""
        if cls._instance is None:
            cls._instance = super().__new__(cls)
        return cls._instance
    
    def __init__(self):
        """초기화"""
        if self._config is None:
            self._config = self._load_config()
    
    def _load_config(self) -> HwpConfig:
        """설정 파일 로드 또는 기본 설정 생성"""
        # 설정 파일 경로 결정
        config_paths = [
            "./hwp_config.json",                    # 현재 디렉토리
            os.path.expanduser("~/.hwp_mcp/config.json"),  # 사용자 홈
            os.path.join(os.path.dirname(__file__), "../../hwp_config.json")  # 프로젝트 루트
        ]
        
        # 환경 변수에서 설정 파일 경로 확인
        env_config_path = os.environ.get("HWP_MCP_CONFIG")
        if env_config_path:
            config_paths.insert(0, env_config_path)
        
        # 설정 파일 찾기
        for path in config_paths:
            if os.path.exists(path):
                try:
                    # 검증된 경로로 로드
                    validated_path = validate_file_path(path, must_exist=True)
                    return HwpConfig.load(validated_path)
                except Exception as e:
                    print(f"설정 파일 로드 실패 ({path}): {e}")
        
        # 기본 설정 반환
        return HwpConfig()
    
    @property
    def config(self) -> HwpConfig:
        """현재 설정 반환"""
        return self._config
    
    def update(self, **kwargs):
        """설정 업데이트"""
        for key, value in kwargs.items():
            if hasattr(self._config, key):
                setattr(self._config, key, value)
    
    def save(self, path: Optional[str] = None):
        """설정 저장"""
        if path is None:
            path = "./hwp_config.json"
        self._config.save(path)
    
    def reset(self):
        """설정 초기화"""
        self._config = HwpConfig()


# 전역 설정 인스턴스
config = ConfigManager().config

# 편의 함수
def get_config() -> HwpConfig:
    """현재 설정 가져오기"""
    return ConfigManager().config

def update_config(**kwargs):
    """설정 업데이트"""
    ConfigManager().update(**kwargs)

def save_config(path: Optional[str] = None):
    """설정 저장"""
    ConfigManager().save(path)

def reset_config():
    """설정 초기화"""
    ConfigManager().reset()