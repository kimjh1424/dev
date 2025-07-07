import sys
import time
from PyQt6.QtWidgets import (
    QApplication, QWidget, QVBoxLayout, QHBoxLayout, 
    QLineEdit, QPushButton, QLabel, QStatusBar, QMainWindow, QMessageBox, QFileDialog,
    QSpinBox
)
from PyQt6.QtCore import Qt, QThread, pyqtSignal
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException, NoSuchElementException, ElementClickInterceptedException
from selenium.webdriver.common.action_chains import ActionChains
import openpyxl
import re

class CrawlerThread(QThread):
    progress_update = pyqtSignal(str)
    finished = pyqtSignal(list)

    def __init__(self, keyword, max_count=10):
        super().__init__()
        self.keyword = keyword
        self.max_count = max_count
        self.is_running = True

    def run(self):
        data = []
        self.progress_update.emit(f"'{self.keyword}' 검색을 시작합니다... (최대 {self.max_count}개 수집)")
        driver = None
        try:
            from urllib.parse import quote
            options = webdriver.ChromeOptions()
            # options.add_argument("--headless")
            options.add_argument("--log-level=3")
            options.add_argument("--disable-blink-features=AutomationControlled")
            options.add_experimental_option("excludeSwitches", ["enable-automation"])
            options.add_experimental_option('useAutomationExtension', False)
            driver = webdriver.Chrome(options=options)
            
            # 1. URL 기반으로 직접 검색 (URL 인코딩)
            encoded_keyword = quote(self.keyword)
            search_url = f"https://map.naver.com/p/search/{encoded_keyword}"
            driver.get(search_url)
            
            # 페이지 로드 대기
            time.sleep(3)

            # 2. searchIframe으로 전환하고, 내부 목록이 로드될 때까지 대기
            WebDriverWait(driver, 20).until(EC.frame_to_be_available_and_switch_to_it((By.ID, "searchIframe")))
            time.sleep(2)  # iframe 전환 후 안정화 대기

            # 3. 스크롤 다운
            try:
                # 더 일반적인 스크롤 컨테이너 선택자 사용
                scroll_container = None
                for selector in ["._3_h-N._1tF3S", "._3_h-N", ".scroll_area"]:
                    try:
                        scroll_container = driver.find_element(By.CSS_SELECTOR, selector)
                        break
                    except NoSuchElementException:
                        continue
                
                if scroll_container:
                    last_height = driver.execute_script("return arguments[0].scrollHeight", scroll_container)
                    scroll_count = 0
                    max_scrolls = 10  # 최대 스크롤 횟수 제한
                    
                    while self.is_running and scroll_count < max_scrolls:
                        self.progress_update.emit(f"결과 목록을 스크롤합니다... ({scroll_count + 1}/{max_scrolls})")
                        driver.execute_script("arguments[0].scrollTo(0, arguments[0].scrollHeight);", scroll_container)
                        time.sleep(2)
                        new_height = driver.execute_script("return arguments[0].scrollHeight", scroll_container)
                        if new_height == last_height:
                            break
                        last_height = new_height
                        scroll_count += 1
                else:
                    self.progress_update.emit("스크롤 컨테이너를 찾을 수 없습니다. 현재 보이는 결과만 수집합니다.")
            except Exception as e:
                self.progress_update.emit(f"스크롤 중 오류: {str(e)}")

            # 4. 모든 장소 링크 찾기 (다양한 선택자 시도)
            time.sleep(2)
            place_elements = []
            
            # 여러 가능한 선택자들을 시도
            selectors = [
                "a.place_bluelink",
                "a[class*='place_bluelink']",
                ".place_bluelink",
                "span.YwYLL",  # 장소명이 있는 span
                "a[role='button']",
                ".VLTHu.OW9LQ"  # 리스트 아이템
            ]
            
            for selector in selectors:
                try:
                    elements = driver.find_elements(By.CSS_SELECTOR, selector)
                    if elements:
                        self.progress_update.emit(f"{len(elements)}개의 요소를 찾았습니다. (선택자: {selector})")
                        # span.YwYLL인 경우 부모 a 태그를 찾아야 함
                        if selector == "span.YwYLL":
                            place_elements = []
                            for elem in elements:
                                try:
                                    parent = elem.find_element(By.XPATH, "./ancestor::a[@role='button']")
                                    place_elements.append(parent)
                                except:
                                    try:
                                        parent = elem.find_element(By.XPATH, "./ancestor::a")
                                        place_elements.append(parent)
                                    except:
                                        pass
                        else:
                            place_elements = elements
                        break
                except NoSuchElementException:
                    continue
            
            if not place_elements:
                self.progress_update.emit("클릭 가능한 장소를 찾을 수 없습니다.")
                return

            # 최대 갯수만큼만 처리
            total_to_process = min(len(place_elements), self.max_count)
            self.progress_update.emit(f"{len(place_elements)}개의 장소를 찾았습니다. {total_to_process}개의 정보를 수집합니다.")

            # 5. 각 장소를 클릭하여 정보 추출
            collected_count = 0
            for i in range(total_to_process):
                if not self.is_running:
                    break
                
                try:
                    # 매번 요소를 다시 찾기 (DOM 변경 대응)
                    driver.switch_to.default_content()
                    WebDriverWait(driver, 10).until(EC.frame_to_be_available_and_switch_to_it((By.ID, "searchIframe")))
                    time.sleep(1)
                    
                    # 현재 인덱스의 요소 다시 찾기
                    current_elements = []
                    for selector in selectors:
                        try:
                            elements = driver.find_elements(By.CSS_SELECTOR, selector)
                            if elements and len(elements) > i:
                                if selector == "span.YwYLL":
                                    try:
                                        parent = elements[i].find_element(By.XPATH, "./ancestor::a[@role='button']")
                                        current_elements = [parent]
                                    except:
                                        try:
                                            parent = elements[i].find_element(By.XPATH, "./ancestor::a")
                                            current_elements = [parent]
                                        except:
                                            continue
                                else:
                                    current_elements = [elements[i]]
                                break
                        except:
                            continue
                    
                    if not current_elements:
                        continue
                    
                    element = current_elements[0]
                    
                    # 요소가 보이도록 스크롤
                    driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", element)
                    time.sleep(1)
                    
                    # 클릭 시도 (여러 방법)
                    clicked = False
                    
                    # 방법 1: JavaScript 클릭
                    try:
                        driver.execute_script("arguments[0].click();", element)
                        clicked = True
                    except:
                        pass
                    
                    # 방법 2: Actions 클릭
                    if not clicked:
                        try:
                            actions = ActionChains(driver)
                            actions.move_to_element(element).click().perform()
                            clicked = True
                        except:
                            pass
                    
                    # 방법 3: 일반 클릭
                    if not clicked:
                        try:
                            element.click()
                            clicked = True
                        except:
                            pass
                    
                    if not clicked:
                        self.progress_update.emit(f"({i + 1}번째 항목 클릭 실패)")
                        continue
                    
                    # 클릭 후 대기
                    time.sleep(2)
                    
                    # entryIframe으로 전환하여 정보 추출
                    driver.switch_to.default_content()
                    
                    try:
                        WebDriverWait(driver, 5).until(EC.frame_to_be_available_and_switch_to_it((By.ID, "entryIframe")))
                        time.sleep(1)
                        
                        # 이름 추출
                        name = "정보 없음"
                        name_selectors = [".YwYLL", ".GHAhO", "span.YwYLL", "h2.YwYLL"]
                        for sel in name_selectors:
                            try:
                                name_elem = WebDriverWait(driver, 3).until(
                                    EC.presence_of_element_located((By.CSS_SELECTOR, sel))
                                )
                                name = name_elem.text.strip()
                                # "복사" 단어 제거
                                name = name.replace('복사', '').strip()
                                if name:
                                    break
                            except:
                                continue
                        
                        # 주소 추출 (도로명과 지번 모두)
                        road_address = "정보 없음"
                        jibun_address = "정보 없음"
                        
                        # 주소 버튼 클릭 시도
                        address_button_clicked = False
                        try:
                            # PkgBl 클래스를 가진 주소 버튼 찾기
                            address_button = driver.find_element(By.CSS_SELECTOR, "a.PkgBl")
                            driver.execute_script("arguments[0].click();", address_button)
                            time.sleep(1)
                            address_button_clicked = True
                        except:
                            pass
                        
                        if address_button_clicked:
                            # 클릭 후 나타나는 주소 정보들 추출
                            try:
                                # 모든 주소 정보를 담고 있는 div들 찾기
                                address_divs = driver.find_elements(By.CSS_SELECTOR, "div.nQ7Lh")
                                
                                for div in address_divs:
                                    try:
                                        # 각 div 내의 텍스트 확인
                                        address_text = div.text.strip()
                                        # "복사" 단어 제거
                                        address_text = address_text.replace('복사', '').strip()
                                        
                                        # 불필요한 레이블 제거
                                        address_text = address_text.replace('도로명', '').replace('지번', '').strip()
                                        
                                        # 도로명 주소 패턴 (도로명이 포함된 주소)
                                        if ('로' in address_text or '길' in address_text) and road_address == "정보 없음":
                                            road_address = address_text
                                        # 지번 주소 패턴 (동/리가 포함되거나 숫자-숫자 형태)
                                        elif ('동' in address_text or '리' in address_text or re.search(r'\d+-\d+', address_text)) and jibun_address == "정보 없음":
                                            jibun_address = address_text
                                    except:
                                        continue
                                
                                # 만약 주소가 하나만 있는 경우
                                if road_address == "정보 없음" and jibun_address == "정보 없음":
                                    try:
                                        # Y31Sf 클래스에서 직접 주소 가져오기
                                        address_container = driver.find_element(By.CSS_SELECTOR, ".Y31Sf")
                                        full_text = address_container.text.strip()
                                        # "복사" 단어 제거
                                        full_text = full_text.replace('복사', '').strip()
                                        # 불필요한 레이블 제거
                                        full_text = full_text.replace('도로명', '').replace('지번', '').strip()
                                        if full_text:
                                            # 첫 번째 줄을 도로명 주소로 사용
                                            lines = full_text.split('\n')
                                            if lines:
                                                road_address = lines[0]
                                    except:
                                        pass
                            except Exception as e:
                                self.progress_update.emit(f"주소 상세 정보 추출 중 오류: {str(e)}")
                        
                        # 주소 버튼 클릭이 안 된 경우 기본 주소 추출 시도
                        if road_address == "정보 없음" and jibun_address == "정보 없음":
                            address_selectors = [".PkgBl", ".Y31Sf", "span.PkgBl", ".address", ".LDgIH"]
                            for sel in address_selectors:
                                try:
                                    address_elem = driver.find_element(By.CSS_SELECTOR, sel)
                                    address_text = address_elem.text.strip()
                                    # "복사" 단어 제거
                                    address_text = address_text.replace('복사', '').strip()
                                    # 불필요한 레이블 제거
                                    address_text = address_text.replace('도로명', '').replace('지번', '').strip()
                                    if address_text:
                                        # 도로명과 지번 구분 시도
                                        if '/' in address_text:
                                            parts = address_text.split('/')
                                            road_address = parts[0].strip()
                                            if len(parts) > 1:
                                                jibun_address = parts[1].strip()
                                        else:
                                            road_address = address_text
                                        break
                                except:
                                    continue
                        
                        # 전화번호 추출
                        phone = "정보 없음"
                        
                        # 전화번호 버튼 클릭 시도
                        phone_button_selectors = [".BfF3H", ".U7pYf", "button[aria-label*='전화']"]
                        for sel in phone_button_selectors:
                            try:
                                phone_button = driver.find_element(By.CSS_SELECTOR, sel)
                                driver.execute_script("arguments[0].click();", phone_button)
                                time.sleep(1)
                                break
                            except:
                                continue
                        
                        # 전화번호 텍스트 찾기
                        phone_selectors = [".J7eF_", ".xlx7Q", ".RiCN3", "span.xlx7Q", "a[href^='tel:']"]
                        for sel in phone_selectors:
                            try:
                                phone_elem = driver.find_element(By.CSS_SELECTOR, sel)
                                phone_text = phone_elem.text.strip()
                                if phone_text and re.search(r'\d{2,}', phone_text):
                                    # "복사" 단어 제거
                                    phone_text = phone_text.replace('복사', '').strip()
                                    # 불필요한 레이블 제거
                                    phone_text = phone_text.replace('휴대전화번호', '').replace('전화번호', '').replace('전화', '').strip()
                                    # 추가적인 정제: 공백 정리
                                    phone_text = ' '.join(phone_text.split())
                                    phone = phone_text
                                    break
                                # href에서 전화번호 추출
                                if sel == "a[href^='tel:']":
                                    href = phone_elem.get_attribute('href')
                                    if href:
                                        phone = href.replace('tel:', '').strip()
                                        break
                            except:
                                continue
                        
                        if name != "정보 없음":
                            # 최종 데이터 정제
                            if road_address != "정보 없음":
                                # 주소 관련 레이블 제거
                                road_address = road_address.replace('도로명', '').replace('복사', '').replace('주소', '').strip()
                                road_address = road_address.strip('[](){}|').strip()
                            if jibun_address != "정보 없음":
                                # 주소 관련 레이블 제거
                                jibun_address = jibun_address.replace('지번', '').replace('복사', '').replace('주소', '').strip()
                                jibun_address = jibun_address.strip('[](){}|').strip()
                            if phone != "정보 없음":
                                # 전화번호 관련 레이블 제거
                                phone = phone.replace('휴대전화번호', '').replace('전화번호', '').replace('전화', '').replace('복사', '').replace('연락처', '').strip()
                                phone = phone.strip('[](){}|').strip()
                                # 전화번호 형식 정리
                                phone = ' '.join(phone.split())
                            
                            data.append([name, road_address, jibun_address, phone])
                            collected_count += 1
                            self.progress_update.emit(f"({collected_count}/{total_to_process}) {name} 정보 수집 완료")
                            
                            # 목표 갯수에 도달하면 중단
                            if collected_count >= self.max_count:
                                self.progress_update.emit(f"목표 수집 갯수({self.max_count}개)에 도달했습니다.")
                                break
                        
                    except TimeoutException:
                        self.progress_update.emit(f"({i + 1}/{total_to_process}번째 항목에서 상세 정보를 찾을 수 없음)")
                    except Exception as e:
                        self.progress_update.emit(f"({i + 1}/{total_to_process}번째 항목 처리 중 오류: {str(e)}")
                    
                except Exception as e:
                    self.progress_update.emit(f"항목 처리 중 오류: {str(e)}")
                    continue

            self.progress_update.emit(f"크롤링 완료. 총 {collected_count}개의 정보를 수집했습니다.")

        except Exception as e:
            self.progress_update.emit(f"크롤링 프로세스 중 오류 발생: {str(e)}")
        finally:
            if driver:
                driver.quit()
            self.finished.emit(data)

    def stop(self):
        self.is_running = False
        self.progress_update.emit("작업을 중단합니다.")


class NaverMapCrawlerApp(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle('Naver Map Crawler v2.0')
        self.setGeometry(100, 100, 700, 280)
        self.crawler_thread = None

        central_widget = QWidget()
        self.setCentralWidget(central_widget)
        main_layout = QVBoxLayout(central_widget)

        # 타이틀
        title_layout = QVBoxLayout()
        title_label = QLabel('네이버 지도 크롤러')
        title_label.setStyleSheet("font-size: 18px; font-weight: bold; padding: 10px;")
        title_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        title_layout.addWidget(title_label)
        main_layout.addLayout(title_layout)

        # 검색 영역
        search_layout = QHBoxLayout()
        self.search_input = QLineEdit()
        self.search_input.setPlaceholderText('검색어를 입력하세요... (예: 강남 카페, 서울 병원)')
        self.search_input.setStyleSheet("padding: 8px; font-size: 14px;")
        
        # 최대 갯수 입력
        count_label = QLabel('최대 갯수:')
        count_label.setStyleSheet("font-size: 14px; padding: 0 5px;")
        self.max_count_input = QSpinBox()
        self.max_count_input.setMinimum(1)
        self.max_count_input.setMaximum(500)
        self.max_count_input.setValue(10)  # 기본값 10개
        self.max_count_input.setStyleSheet("padding: 8px; font-size: 14px; width: 80px;")
        self.max_count_input.setToolTip("수집할 최대 정보 갯수 (1~500)")
        
        self.search_button = QPushButton('검색 시작')
        self.search_button.setStyleSheet("padding: 8px 16px; font-size: 14px;")
        self.search_button.clicked.connect(self.start_crawling)
        
        # Enter 키로도 검색 가능
        self.search_input.returnPressed.connect(self.start_crawling)
        
        search_layout.addWidget(self.search_input)
        search_layout.addWidget(count_label)
        search_layout.addWidget(self.max_count_input)
        search_layout.addWidget(self.search_button)
        main_layout.addLayout(search_layout)
        
        # 안내 메시지
        info_label = QLabel('• 검색 결과를 엑셀 파일로 저장합니다.\n• 수집 항목: 장소명, 도로명 주소, 지번 주소, 전화번호\n• 최대 갯수를 조절하여 원하는 만큼만 수집할 수 있습니다.')
        info_label.setStyleSheet("padding: 10px; color: #666;")
        main_layout.addWidget(info_label)
        
        main_layout.addStretch(1)
        
        # 하단 정보
        bottom_layout = QHBoxLayout()
        by_label = QLabel('By ANYCODER | v2.0')
        by_label.setStyleSheet("color: #888; font-size: 12px;")
        bottom_layout.addStretch(1)
        bottom_layout.addWidget(by_label)
        main_layout.addLayout(bottom_layout)

        # 상태바
        self.status_bar = QStatusBar()
        self.setStatusBar(self.status_bar)
        self.status_bar.showMessage('준비 완료')

    def start_crawling(self):
        keyword = self.search_input.text().strip()
        if not keyword:
            QMessageBox.warning(self, "입력 오류", "검색어를 입력해주세요.")
            return

        max_count = self.max_count_input.value()
        
        self.search_button.setEnabled(False)
        self.search_button.setText("크롤링 중...")
        self.search_input.setEnabled(False)
        self.max_count_input.setEnabled(False)
        
        self.crawler_thread = CrawlerThread(keyword, max_count)
        self.crawler_thread.progress_update.connect(self.update_status)
        self.crawler_thread.finished.connect(self.crawling_finished)
        self.crawler_thread.start()

    def update_status(self, message):
        self.status_bar.showMessage(message)

    def crawling_finished(self, data):
        self.search_button.setEnabled(True)
        self.search_button.setText("검색 시작")
        self.search_input.setEnabled(True)
        self.max_count_input.setEnabled(True)
        
        if data:
            self.status_bar.showMessage(f"크롤링 완료. {len(data)}개의 정보를 수집했습니다.")
            self.save_to_excel(data)
        else:
            QMessageBox.information(self, "결과 없음", "수집된 데이터가 없습니다.\n검색어를 확인하고 다시 시도해주세요.")
            self.status_bar.showMessage("데이터 수집 실패")

    def save_to_excel(self, data):
        if not data:
            return
        
        # 기본 파일명 생성
        keyword = self.search_input.text().strip()
        default_filename = f"네이버지도_{keyword}_{time.strftime('%Y%m%d_%H%M%S')}.xlsx"
        
        file_path, _ = QFileDialog.getSaveFileName(
            self, 
            "엑셀 파일 저장", 
            default_filename, 
            "Excel Files (*.xlsx)"
        )
        
        if file_path:
            try:
                workbook = openpyxl.Workbook()
                sheet = workbook.active
                sheet.title = "네이버 지도 크롤링 결과"
                
                # 헤더 추가 및 스타일링
                headers = ["번호", "장소명", "도로명 주소", "지번 주소", "전화번호"]
                sheet.append(headers)
                
                # 헤더 스타일
                for cell in sheet[1]:
                    cell.font = openpyxl.styles.Font(bold=True)
                    cell.fill = openpyxl.styles.PatternFill(
                        start_color="DDDDDD", 
                        end_color="DDDDDD", 
                        fill_type="solid"
                    )
                
                # 데이터 추가
                for idx, item in enumerate(data, 1):
                    sheet.append([idx] + item)
                
                # 열 너비 자동 조정
                from openpyxl.worksheet.merge import MergedCell
                
                for col_idx, column in enumerate(sheet.columns, 1):
                    max_length = 0
                    column_letter = openpyxl.utils.get_column_letter(col_idx)
                    
                    for cell in column:
                        try:
                            # MergedCell이 아닌 경우에만 처리
                            if not isinstance(cell, MergedCell) and cell.value:
                                if len(str(cell.value)) > max_length:
                                    max_length = len(str(cell.value))
                        except:
                            pass
                    
                    adjusted_width = min(max_length + 2, 50)
                    if adjusted_width > 0:
                        sheet.column_dimensions[column_letter].width = adjusted_width
                    
                workbook.save(file_path)
                
                reply = QMessageBox.information(
                    self, 
                    "저장 완료", 
                    f"'{file_path}'에 성공적으로 저장되었습니다.\n\n수집된 데이터: {len(data)}개",
                    QMessageBox.StandardButton.Ok
                )
                self.status_bar.showMessage(f"저장 완료: {file_path}")
                
            except Exception as e:
                QMessageBox.critical(self, "저장 실패", f"엑셀 파일 저장 중 오류가 발생했습니다:\n{str(e)}")
                self.status_bar.showMessage("저장 실패")

    def closeEvent(self, event):
        if self.crawler_thread and self.crawler_thread.isRunning():
            reply = QMessageBox.question(
                self, 
                '작업 중', 
                "크롤링 작업이 아직 진행 중입니다.\n종료하시겠습니까?",
                QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No,
                QMessageBox.StandardButton.No
            )

            if reply == QMessageBox.StandardButton.Yes:
                self.crawler_thread.stop()
                self.crawler_thread.wait()
                event.accept()
            else:
                event.ignore()
        else:
            event.accept()


if __name__ == '__main__':
    app = QApplication(sys.argv)
    app.setStyle('Fusion')
    ex = NaverMapCrawlerApp()
    ex.show()
    sys.exit(app.exec())
