# NERPpkb

## 개요
* SAP 제어 편의성을 위해 개발한 패키지
* 현재 repository를 포함한 모든 공개 repository의 데이터는 전부 Dummy데이터로, 사용 전 교체하여 사용 필요

## 필수사항
- SAPGUI : SAP Scripting 활용
- SAP권한 : 기능별 필요한 권한 모두 미리 받아두어야 함
- 사용 전 실제 데이터로 바꿔두고 사용해야함 (현재 Dummy data임)

## 파일별 설명
* NERPpkb : 모든 기능은 이 파일에 통합되어있음
  * check_and_open_sap(server_name:str, id:str, pw:str, windows:int=3) : 현재 상태 확인 후 SAP 실행
  * sap_login(guiApp, server_name:str, id:str, pw:str) : 위 함수에서 실제 로그인시 사용
  * start_menu_with_tcode(session, tcode:str) : 현재 접속상태 확인 후 지정한 T-CODE접속
  * loop_tcode(session, tcode:str) : 입력한 T-CODE의 실행/종료 반복(프로세스 유지용)
  * error_handler_pi(session, error_txt:str) : 입력한 메시지의 팝업에 대해 끄고 다음 내용 진행
  * chk_exist_pi_lc(session, pi_name:str) : 현재 입력하고자 하는 pi_name이 이미 있는 건인지 조회
  * input_pi_lc(session, lc_org:int, pi_info:list[str], date:list[str], main_info:list[str,bool], port_and_address_txt:list[str], is_local:bool=False) : 입력한 값을 바탕으로 PI등록
  * nego_history_download(session, local_select:str, nego_criteria:list[str], down_path:str) : 네고History 다운로드용
  * pouch_download(session, pouch_criteria:list[str], down_path:str) : PouchHistory 다운로드용
  * down_billing_detail(session, companycd:str, date_start:str, date_end:str, down_path:str, org_codes:list[str]) : 선적History 다운로드용

* NERP_login : 로그인 기능만 따로 사용 (프로세스 유지기능 포함)
  * Jupyter notebook으로, 셀 실행시 바로 사용 가능

* NERP_downloads : 실적 등 엑셀로 받고자 할 때 사용
  * Jupyter notebook으로, 셀 실행시 바로 사용 가능