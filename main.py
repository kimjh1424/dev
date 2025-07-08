import sys
import time
import tkinter as tk
from tkinter import ttk, messagebox, filedialog
import threading
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException, NoSuchElementException, ElementClickInterceptedException, StaleElementReferenceException
from selenium.webdriver.common.action_chains import ActionChains
import openpyxl
import re
from urllib.parse import quote

class CrawlerThread(threading.Thread):
    def __init__(self, keyword, max_count, callback, status_callback, headless_mode=False):
        super().__init__()
        self.keyword = keyword
        self.max_count = max_count
        self.callback = callback
        self.status_callback = status_callback
        self.is_running = True
        self.daemon = True
        self.headless_mode = headless_mode

    def find_next_page_button(self, driver):
        """다음 페이지 버튼을 정확히 찾는 함수"""
        selectors = ["a.eUTV2", "a._2PoiJ", "a[class*='page']", "button[class*='next']"]
        
        for selector in selectors:
            try:
                buttons = driver.find_elements(By.CSS_SELECTOR, selector)
                
                for button in buttons:
                    if button.get_attribute("aria-disabled") == "true":
                        continue
                    
                    try:
                        svg_element = button.find_element(By.TAG_NAME, "svg")
                        path_element = svg_element.find_element(By.TAG_NAME, "path")
                        d_attribute = path_element.get_attribute("d")
                        
                        if d_attribute and any(pattern in d_attribute for pattern in ["M12", "M14", "M10.524"]):
                            if not any(pattern in d_attribute for pattern in ["M8", "M6", "M13.476"]):
                                self.status_callback(f"다음 페이지 버튼 찾음: {selector}")
                                return button
                                
                    except:
                        button_text = button.text.strip()
                        if "다음" in button_text or ">" in button_text:
                            return button
            except:
                continue
                
        return None

    def turbo_scroll_to_load_all(self, driver, scroll_container):
        """터보 스크롤 - 확실하게 70개 모두 로드"""
        self.status_callback("⚡ 스크롤 컨테이너 찾는 중...")
        
        # 네이버 지도 전용 스크롤 컨테이너 찾기
        scroll_found = False
        scroll_container = None
        
        # 방법 1: ID로 직접 찾기
        try:
            scroll_container = driver.find_element(By.ID, "_list_scroll_container")
            self.status_callback("✅ _list_scroll_container 찾음!")
            scroll_found = True
        except:
            pass
        
        # 방법 2: 클래스로 찾기
        if not scroll_found:
            try:
                scroll_container = driver.find_element(By.CSS_SELECTOR, "div._2ky45")
                self.status_callback("✅ div._2ky45 찾음!")
                scroll_found = True
            except:
                pass
        
        # 방법 3: JavaScript로 스크롤 가능한 요소 찾기
        if not scroll_found:
            try:
                scroll_container = driver.execute_script("""
                    // 모든 div 요소 확인
                    var divs = document.querySelectorAll('div');
                    for (var i = 0; i < divs.length; i++) {
                        if (divs[i].scrollHeight > divs[i].clientHeight && 
                            divs[i].scrollHeight > 500) {
                            console.log('스크롤 가능한 요소 찾음:', divs[i]);
                            return divs[i];
                        }
                    }
                    return null;
                """)
                if scroll_container:
                    self.status_callback("✅ JavaScript로 스크롤 컨테이너 찾음!")
                    scroll_found = True
            except:
                pass
        
        # 방법 4: iframe 내부 body
        if not scroll_found:
            try:
                scroll_container = driver.find_element(By.TAG_NAME, "body")
                self.status_callback("⚠️ body를 스크롤 컨테이너로 사용")
            except:
                self.status_callback("❌ 스크롤 컨테이너를 찾을 수 없음!")
                return 0
        
        # 초기 아이템 수 확인
        initial_items = driver.find_elements(By.CSS_SELECTOR, "a.place_bluelink")
        if not initial_items:
            initial_items = driver.find_elements(By.CSS_SELECTOR, "li.UEzoS")
        
        previous_count = len(initial_items)
        self.status_callback(f"초기 아이템: {previous_count}개")
        
        # 스크롤 전략 1: 페이지 끝까지 여러번 스크롤
        self.status_callback("⚡ 끝까지 스크롤 시작...")
        
        for i in range(10):  # 최대 10번 시도
            try:
                # JavaScript로 강제 스크롤
                driver.execute_script("""
                    var element = arguments[0];
                    element.scrollTop = element.scrollHeight;
                """, scroll_container)
                
                time.sleep(0.5)  # 로드 대기
                
                # 현재 아이템 수 확인
                current_items = driver.find_elements(By.CSS_SELECTOR, "a.place_bluelink")
                if not current_items:
                    current_items = driver.find_elements(By.CSS_SELECTOR, "li.UEzoS")
                
                current_count = len(current_items)
                
                if current_count > previous_count:
                    self.status_callback(f"⚡ {current_count}개 로드됨 (+{current_count - previous_count})")
                    previous_count = current_count
                    
                    # 70개 도달시 종료
                    if current_count >= 70:
                        self.status_callback(f"✅ 70개 도달! (현재: {current_count}개)")
                        break
                elif i > 3:  # 3번 이상 변화없으면 종료
                    self.status_callback(f"⚠️ 더 이상 로드되지 않음 (최종: {current_count}개)")
                    break
                    
            except Exception as e:
                self.status_callback(f"스크롤 오류: {e}")
        
        # 스크롤 전략 2: 단계적 스크롤 (위 방법이 실패시)
        if previous_count < 70:
            self.status_callback("⚡ 단계적 스크롤 시도...")
            
            try:
                # 전체 높이 가져오기
                total_height = driver.execute_script("return arguments[0].scrollHeight", scroll_container)
                step = total_height // 5  # 5단계로 나누어 스크롤
                
                for position in range(step, total_height + step, step):
                    driver.execute_script(f"""
                        arguments[0].scrollTop = {position};
                    """, scroll_container)
                    time.sleep(0.3)
                    
                    # 아이템 수 확인
                    current_items = driver.find_elements(By.CSS_SELECTOR, "a.place_bluelink")
                    if not current_items:
                        current_items = driver.find_elements(By.CSS_SELECTOR, "li.UEzoS")
                    
                    if len(current_items) >= 70:
                        self.status_callback(f"✅ 70개 도달! (현재: {len(current_items)}개)")
                        break
                        
            except Exception as e:
                self.status_callback(f"단계적 스크롤 오류: {e}")
        
        # 스크롤 전략 3: 키보드 이벤트 사용
        if previous_count < 70:
            self.status_callback("⚡ 키보드 스크롤 시도...")
            
            try:
                # 스크롤 컨테이너에 포커스
                driver.execute_script("arguments[0].focus();", scroll_container)
                
                # End 키를 여러번 누르기
                from selenium.webdriver.common.keys import Keys
                actions = ActionChains(driver)
                
                for _ in range(10):
                    actions.send_keys(Keys.END).perform()
                    time.sleep(0.3)
                    
                    # 아이템 수 확인
                    current_items = driver.find_elements(By.CSS_SELECTOR, "a.place_bluelink")
                    if not current_items:
                        current_items = driver.find_elements(By.CSS_SELECTOR, "li.UEzoS")
                    
                    if len(current_items) >= 70:
                        self.status_callback(f"✅ 70개 도달! (현재: {len(current_items)}개)")
                        break
                        
            except Exception as e:
                self.status_callback(f"키보드 스크롤 오류: {e}")
        
        # 맨 위로 스크롤
        try:
            driver.execute_script("arguments[0].scrollTop = 0", scroll_container)
        except:
            driver.execute_script("window.scrollTo(0, 0);")
        
        time.sleep(0.5)
        
        # 최종 아이템 수 확인
        final_items = driver.find_elements(By.CSS_SELECTOR, "a.place_bluelink")
        if not final_items:
            final_items = driver.find_elements(By.CSS_SELECTOR, "li.UEzoS")
        
        self.status_callback(f"⚡ 스크롤 완료! 최종 로드: {len(final_items)}개")
        
        if len(final_items) < 70 and len(final_items) > 0:
            self.status_callback(f"⚠️ 70개 미만 로드됨. 해당 검색어의 결과가 {len(final_items)}개일 수 있습니다.")
        
        return len(final_items)

    def run(self):
        data = []
        collected_count = 0
        self.status_callback(f"⚡ '{self.keyword}' 터보 크롤링 시작! (목표: {self.max_count}개)")
        if self.max_count > 100:
            self.status_callback("🚀 대량 크롤링 모드 활성화! 메모리 최적화 적용")
        if self.headless_mode:
            self.status_callback("👻 헤드리스 모드로 실행중...")
        else:
            self.status_callback("👀 일반 모드로 실행중 (브라우저 표시)")
        self.status_callback("=" * 50)
        driver = None
        search_url = None

        try:
            options = webdriver.ChromeOptions()
            
            # 헤드리스 모드 설정
            if self.headless_mode:
                options.add_argument("--headless")
                self.status_callback("헤드리스 모드 활성화")
            
            options.add_argument("--log-level=3")
            options.add_argument("--disable-blink-features=AutomationControlled")
            options.add_experimental_option("excludeSwitches", ["enable-automation"])
            options.add_experimental_option('useAutomationExtension', False)
            
            # 안정성 개선 옵션
            options.add_argument("--no-sandbox")
            options.add_argument("--disable-dev-shm-usage")
            options.add_argument("--disable-gpu")
            options.add_argument("--disable-features=VizDisplayCompositor")
            
            # 메모리 관리 옵션 추가
            options.add_argument("--max_old_space_size=4096")
            options.add_argument("--memory-pressure-off")
            options.add_argument("--disable-background-timer-throttling")
            options.add_argument("--disable-renderer-backgrounding")
            options.add_argument("--disable-features=TranslateUI")
            options.add_argument("--disable-ipc-flooding-protection")
            
            # 대량 크롤링을 위한 추가 옵션
            if self.max_count > 100:
                options.add_argument("--disable-logging")
                options.add_argument("--disable-gpu-sandbox")
                options.add_argument("--disable-software-rasterizer")
                options.add_argument("--disable-background-timer-throttling")
                options.add_argument("--disable-backgrounding-occluded-windows")
                options.add_argument("--disable-renderer-backgrounding")
                options.add_argument("--disable-features=TranslateUI")
                options.add_argument("--disable-ipc-flooding-protection")
                # Chrome 메모리 사용량 제한
                options.add_argument("--max_old_space_size=8192")  # 8GB까지 허용
            
            # 속도 개선 옵션
            options.add_argument("--disable-images")  # 이미지 로드 안함
            options.page_load_strategy = 'eager'
            
            driver = webdriver.Chrome(options=options)
            driver.set_page_load_timeout(20)
            driver.implicitly_wait(5)
            
            encoded_keyword = quote(self.keyword)
            search_url = f"https://map.naver.com/p/search/{encoded_keyword}"
            self.status_callback(f"검색 URL: {search_url}")
            driver.get(search_url)
            time.sleep(2)  # 초기 로드 대기

            self.status_callback("searchIframe으로 전환...")
            WebDriverWait(driver, 10).until(EC.frame_to_be_available_and_switch_to_it((By.ID, "searchIframe")))
            time.sleep(1)
            self.status_callback("✅ searchIframe 전환 성공")
            
            # 스크롤 컨테이너 정보 디버깅
            try:
                debug_info = driver.execute_script("""
                    var container = document.getElementById('_list_scroll_container');
                    if (container) {
                        return {
                            found: true,
                            scrollHeight: container.scrollHeight,
                            clientHeight: container.clientHeight,
                            scrollable: container.scrollHeight > container.clientHeight
                        };
                    }
                    return {found: false};
                """)
                self.status_callback(f"📊 스크롤 컨테이너 정보: {debug_info}")
            except:
                pass
            
            # 초기 페이지 로드 확인
            try:
                WebDriverWait(driver, 10).until(
                    EC.presence_of_element_located((By.CSS_SELECTOR, "a.place_bluelink, li.UEzoS"))
                )
                self.status_callback("✅ 검색 결과 로드 확인")
            except:
                self.status_callback("❌ 검색 결과를 찾을 수 없습니다.")
                return

            current_page = 1

            while self.is_running and collected_count < self.max_count:
                self.status_callback(f"\n{'='*50}")
                self.status_callback(f"📌 {current_page} 페이지 터보 크롤링 시작")
                self.status_callback(f"📌 현재까지 수집: {collected_count}개")
                self.status_callback(f"{'='*50}")

                # 터보 스크롤로 모든 아이템 로드 (70개까지)
                loaded_count = self.turbo_scroll_to_load_all(driver, None)
                
                # 모든 장소 요소 가져오기
                all_place_elements = []
                place_selectors = ["a.place_bluelink", "li.UEzoS.rTjJo", "li.UEzoS"]
                
                for selector in place_selectors:
                    try:
                        all_place_elements = driver.find_elements(By.CSS_SELECTOR, selector)
                        if all_place_elements:
                            self.status_callback(f"장소 요소 찾음: {selector} ({len(all_place_elements)}개)")
                            break
                    except:
                        continue

                if not all_place_elements:
                    self.status_callback(f"❌ {current_page} 페이지에서 장소를 찾을 수 없습니다.")
                    break

                page_target_count = len(all_place_elements)
                page_collected = 0
                self.status_callback(f"⚡ {page_target_count}개 장소 크롤링 시작!")

                # 각 장소 크롤링
                for i, element in enumerate(all_place_elements):
                    if not self.is_running or collected_count >= self.max_count:
                        break

                    place_name_for_log = "알 수 없는 장소"
                    try:
                        # 요소 클릭
                        if element.tag_name == 'a':
                            element_to_click = element
                        else:
                            try:
                                element_to_click = element.find_element(By.CSS_SELECTOR, "a.place_bluelink")
                            except:
                                element_to_click = element.find_element(By.CSS_SELECTOR, "a")

                        # JavaScript 클릭 (더 빠름)
                        driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", element_to_click)
                        time.sleep(0.2)  # 0.3에서 0.2로 단축
                        driver.execute_script("arguments[0].click();", element_to_click)
                        time.sleep(1)  # 1.5에서 1초로 단축

                        driver.switch_to.default_content()
                        WebDriverWait(driver, 5).until(EC.frame_to_be_available_and_switch_to_it((By.ID, "entryIframe")))
                        time.sleep(0.5)

                        # 데이터 추출 (JavaScript로 빠르게 + 지번주소 포함)
                        place_data = driver.execute_script("""
                            var result = {
                                name: '',
                                roadAddr: '', 
                                jibunAddr: '',
                                phone: ''
                            };
                            
                            // 이름
                            var nameElem = document.querySelector('.YwYLL, ._3Apjo, .GHAhO, h2');
                            if (nameElem) result.name = nameElem.textContent.replace('복사', '').trim();
                            
                            // 도로명 주소
                            var roadElem = document.querySelector('.LDgIH');
                            if (roadElem) result.roadAddr = roadElem.textContent.replace('복사', '').trim();
                            
                            // 전화번호
                            var phoneElem = document.querySelector('.xlx7Q, ._3ZA58 span, .dry01, .J7eF_');
                            if (phoneElem) {
                                result.phone = phoneElem.textContent
                                    .replace('휴대전화번호', '')
                                    .replace('복사', '')
                                    .trim();
                            }
                            
                            return result;
                        """)

                        name = place_data.get('name', '정보 없음')
                        road_address = place_data.get('roadAddr', '정보 없음')
                        jibun_address = "정보 없음"
                        phone = place_data.get('phone', '정보 없음')
                        
                        # 지번주소 가져오기 - PkgBl 버튼 클릭
                        try:
                            # PkgBl 버튼 찾기 및 클릭
                            address_button = driver.find_element(By.CSS_SELECTOR, "a.PkgBl")
                            driver.execute_script("arguments[0].click();", address_button)
                            time.sleep(0.3)
                            
                            # 지번주소 찾기 (여러 방법 시도)
                            jibun_found = False
                            
                            # 방법 1: nQ7Lh 클래스에서 지번 찾기
                            try:
                                time.sleep(0.2)  # 메뉴 열리는 시간 대기
                                address_items = driver.find_elements(By.CSS_SELECTOR, ".nQ7Lh")
                                for item in address_items:
                                    item_text = item.text.strip()
                                    if "지번" in item_text:
                                        jibun_address = item_text.replace("지번", "").replace("복사", "").strip()
                                        jibun_found = True
                                        self.status_callback(f"📍 지번주소 발견: {jibun_address[:20]}...")
                                        break
                            except:
                                pass
                            
                            # 방법 2: aria-label로 찾기
                            if not jibun_found:
                                try:
                                    jibun_links = driver.find_elements(By.CSS_SELECTOR, "a[aria-label*='지번']")
                                    if jibun_links:
                                        jibun_text = jibun_links[0].text.replace("복사", "").strip()
                                        if jibun_text:
                                            jibun_address = jibun_text
                                            jibun_found = True
                                except:
                                    pass
                            
                            # 방법 3: JavaScript로 모든 텍스트에서 지번 찾기
                            if not jibun_found:
                                jibun_js = driver.execute_script("""
                                    // 모든 요소에서 지번 텍스트 찾기
                                    var allElements = document.querySelectorAll('*');
                                    for (var i = 0; i < allElements.length; i++) {
                                        var text = allElements[i].textContent;
                                        if (text && text.includes('지번') && !text.includes('도로명')) {
                                            // 지번 다음의 텍스트 추출
                                            var match = text.match(/지번\\s*([^복사]+)/);
                                            if (match) {
                                                return match[1].trim();
                                            }
                                        }
                                    }
                                    
                                    // 대안: _2yqUQ 클래스의 두 번째 요소
                                    var addrs = document.querySelectorAll('._2yqUQ');
                                    if (addrs.length > 1) {
                                        return addrs[1].textContent.replace('복사', '').trim();
                                    }
                                    
                                    // 대안2: 도로명 주소가 아닌 다른 주소 찾기
                                    var roadAddr = document.querySelector('.LDgIH').textContent;
                                    var allAddrs = document.querySelectorAll('span');
                                    for (var j = 0; j < allAddrs.length; j++) {
                                        var addr = allAddrs[j].textContent.trim();
                                        if (addr && addr.length > 5 && !addr.includes('복사') && 
                                            !addr.includes(roadAddr) && /[가-힣]/.test(addr)) {
                                            return addr;
                                        }
                                    }
                                    
                                    return null;
                                """)
                                if jibun_js:
                                    jibun_address = jibun_js
                                    jibun_found = True
                            
                            # 버튼 다시 클릭해서 닫기
                            if jibun_found:
                                try:
                                    # 주소 메뉴 닫기
                                    close_button = driver.find_element(By.CSS_SELECTOR, "a.PkgBl[aria-expanded='true']")
                                    driver.execute_script("arguments[0].click();", close_button)
                                except:
                                    pass
                                
                        except Exception as e:
                            # PkgBl 버튼이 없는 경우 대체 방법
                            try:
                                # 여러 주소 요소 찾기
                                address_elems = driver.find_elements(By.CSS_SELECTOR, "._2yqUQ, .LDgIH")
                                if len(address_elems) > 1:
                                    jibun_address = address_elems[1].text.replace("복사", "").strip()
                            except:
                                pass

                        # 전화번호 버튼 클릭 시도
                        if phone == "정보 없음" or not any(char.isdigit() for char in phone):
                            try:
                                phone_button = driver.find_element(By.CSS_SELECTOR, "a.BfF3H")
                                if "전화번호 보기" in phone_button.text:
                                    driver.execute_script("arguments[0].click();", phone_button)
                                    time.sleep(0.3)
                                    phone_elem = driver.find_element(By.CSS_SELECTOR, ".J7eF_")
                                    phone = phone_elem.text.replace("휴대전화번호", "").replace("복사", "").strip()
                            except:
                                pass

                        # 장소명
                        if name != "정보 없음":
                            data.append([name, road_address, jibun_address, phone])
                            collected_count += 1
                            page_collected += 1
                            
                            # 진행 상황 로그 (10개마다)
                            if collected_count % 10 == 0:
                                self.status_callback(f"✅ ({collected_count}/{self.max_count}) 수집 진행중...")
                            else:
                                # 디버그 모드가 아니면 개별 로그 생략
                                pass
                            
                            # 세션 정리 (50개마다 메모리 관리)
                            if collected_count % 50 == 0:
                                try:
                                    driver.execute_script("window.dispatchEvent(new Event('beforeunload'));")
                                except:
                                    pass
                        else:
                            page_collected += 1

                    except Exception as e:
                        page_collected += 1
                        
                        # 탭 크래시 감지
                        if "tab crashed" in str(e).lower() or "session" in str(e).lower():
                            self.status_callback("🛑 Chrome 탭 크래시 감지!")
                            self.status_callback(f"💾 현재까지 수집된 데이터: {collected_count}개")
                            self.is_running = False
                            break
                        continue
                    finally:
                        try:
                            driver.switch_to.default_content()
                            WebDriverWait(driver, 5).until(EC.frame_to_be_available_and_switch_to_it((By.ID, "searchIframe")))
                            time.sleep(0.3)
                        except:
                            # iframe 전환 실패시 계속 진행
                            pass

                self.status_callback(f"⚡ {current_page} 페이지 완료! ({page_collected}/{page_target_count}개 수집)")
                
                # 크래시로 인한 종료 체크
                if not self.is_running:
                    self.status_callback("⚠️ 비정상 종료 감지")
                    break
                
                # 메모리 정리 (20개마다)
                if collected_count > 0 and collected_count % 20 == 0:
                    self.status_callback("💾 메모리 정리 중...")
                    try:
                        # JavaScript 메모리 정리
                        driver.execute_script("""
                            if(typeof gc === 'function') { gc(); }
                            // DOM 정리
                            document.body.style.display = 'none';
                            document.body.offsetHeight;
                            document.body.style.display = '';
                        """)
                        
                        # 쿠키 및 로컬 스토리지 정리
                        driver.delete_all_cookies()
                        driver.execute_script("window.localStorage.clear();")
                        driver.execute_script("window.sessionStorage.clear();")
                        
                        time.sleep(0.5)
                    except:
                        pass
                
                # 50개마다 iframe 재로드
                if collected_count > 0 and collected_count % 50 == 0:
                    self.status_callback("🔄 iframe 새로고침...")
                    try:
                        driver.switch_to.default_content()
                        driver.refresh()
                        time.sleep(2)
                        WebDriverWait(driver, 10).until(EC.frame_to_be_available_and_switch_to_it((By.ID, "searchIframe")))
                        time.sleep(1)
                    except:
                        pass
                
                # 100개 도달시 경고 및 중단 옵션
                if collected_count >= 100 and collected_count < 110:
                    self.status_callback("⚠️ 100개 도달! 안정성을 위해 곧 중단됩니다...")
                
                # 110개에서 자동 중단 (크래시 방지)
                if collected_count >= 110:
                    self.status_callback("🛑 110개 도달! 크래시 방지를 위해 중단합니다.")
                    self.status_callback("💡 나머지는 새로운 검색으로 계속하세요.")
                    break
                
                if collected_count >= self.max_count:
                    self.status_callback(f"✅ 목표 달성! {collected_count}개 수집 완료!")
                    break

                # 다음 페이지로 이동
                if page_collected >= page_target_count * 0.8:  # 80% 이상 수집시 다음 페이지
                    try:
                        driver.switch_to.default_content()
                        WebDriverWait(driver, 5).until(EC.frame_to_be_available_and_switch_to_it((By.ID, "searchIframe")))
                        
                        next_button = self.find_next_page_button(driver)
                        if next_button:
                            driver.execute_script("arguments[0].click();", next_button)
                            self.status_callback(f"⚡ 다음 페이지로 이동!")
                            current_page += 1
                            time.sleep(1.5)
                        else:
                            self.status_callback("❌ 마지막 페이지입니다.")
                            break
                            
                    except Exception as e:
                        self.status_callback(f"❌ 페이지 이동 실패: {e}")
                        break
                else:
                    self.status_callback("❌ 수집률이 낮아 종료합니다.")
                    break

            self.status_callback(f"\n🎉 터보 크롤링 완료!\n총 {collected_count}개 수집 ({current_page}페이지)")

        except Exception as e:
            error_msg = str(e)
            if "tab crashed" in error_msg.lower():
                self.status_callback("🛑 Chrome 탭 크래시! 메모리 부족으로 인한 중단")
                self.status_callback(f"💡 팁: 100개 이하로 설정하거나 헤드리스 모드를 사용하세요")
            else:
                self.status_callback(f"❌ 크롤링 오류: {error_msg[:100]}...")  # 에러 메시지 일부만 표시
        finally:
            if driver:
                driver.quit()
            self.status_callback(f"최종 수집 데이터: {len(data)}개")
            if self.root:
                self.root.after(0, self.callback, data)
            else:
                self.callback(data)

    def stop(self):
        self.is_running = False


class NaverMapCrawlerApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Naver Map Turbo Crawler v3.3 - 대량 크롤링 지원")
        self.root.geometry("800x650")
        
        # 스타일 설정
        style = ttk.Style()
        style.configure('Title.TLabel', font=('Arial', 18, 'bold'))
        style.configure('Turbo.TLabel', font=('Arial', 12, 'bold'), foreground='#ff6b00')
        
        self.crawler_thread = None
        self.setup_ui()
        
    def setup_ui(self):
        # 메인 프레임
        main_frame = ttk.Frame(self.root, padding="10")
        main_frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        
        # 윈도우 크기 조정
        self.root.grid_rowconfigure(0, weight=1)
        self.root.grid_columnconfigure(0, weight=1)
        main_frame.grid_rowconfigure(6, weight=1)  # 로그 프레임 row 조정
        
        # 타이틀
        title_label = ttk.Label(main_frame, text="⚡ 네이버 지도 터보 크롤러 v3.3", style='Title.TLabel')
        title_label.grid(row=0, column=0, columnspan=3, pady=(0, 5))
        
        # 부제목
        subtitle_label = ttk.Label(main_frame, text="500개+ 대량 크롤링 & 지번주소 지원", style='Turbo.TLabel')
        subtitle_label.grid(row=1, column=0, columnspan=3, pady=(0, 10))
        
        # 검색 영역
        search_frame = ttk.Frame(main_frame)
        search_frame.grid(row=2, column=0, columnspan=3, pady=5)
        
        ttk.Label(search_frame, text="검색어:").grid(row=0, column=0, padx=5)
        self.search_entry = ttk.Entry(search_frame, width=30)
        self.search_entry.grid(row=0, column=1, padx=5)
        self.search_entry.bind('<Return>', lambda e: self.start_crawling())
        
        ttk.Label(search_frame, text="최대 갯수:").grid(row=0, column=2, padx=5)
        self.max_count_var = tk.StringVar(value="100")
        self.max_count_spinbox = ttk.Spinbox(search_frame, from_=1, to=1000, textvariable=self.max_count_var, width=10)
        self.max_count_spinbox.grid(row=0, column=3, padx=5)
        
        self.search_button = ttk.Button(search_frame, text="⚡ 터보 시작", command=self.start_crawling)
        self.search_button.grid(row=0, column=4, padx=10)
        
        # 옵션 프레임
        option_frame = ttk.LabelFrame(main_frame, text="⚙️ 터보 옵션", padding="5")
        option_frame.grid(row=3, column=0, columnspan=3, pady=10, sticky=(tk.W, tk.E))
        
        # 헤드리스 모드 체크박스
        self.headless_var = tk.BooleanVar(value=False)
        self.headless_checkbox = ttk.Checkbutton(
            option_frame,
            text="👻 헤드리스 모드 (브라우저 숨김 - 더 빠름)",
            variable=self.headless_var
        )
        self.headless_checkbox.grid(row=0, column=0, padx=10, pady=5)
        
        # 모드 설명
        mode_info = tk.Label(
            option_frame, 
            text="• 일반 모드: 브라우저 표시 (진행 상황 확인 가능)\n• 헤드리스 모드: 브라우저 숨김 (더 빠른 속도)",
            justify=tk.LEFT,
            fg='gray'
        )
        mode_info.grid(row=1, column=0, padx=10, pady=5)
        
        # 안내 메시지
        info_text = "• 터보 스크롤로 빠른 수집 (최대 1000개)\n• 수집 항목: 장소명, 도로명 주소, 지번 주소, 전화번호\n• 100개마다 메모리 자동 정리"
        info_label = ttk.Label(main_frame, text=info_text, foreground='gray')
        info_label.grid(row=5, column=0, columnspan=3, pady=5)
        
        # 로그 프레임
        log_frame = ttk.LabelFrame(main_frame, text="⚡ 터보 로그", padding="5")
        log_frame.grid(row=6, column=0, columnspan=3, sticky=(tk.W, tk.E, tk.N, tk.S), pady=10)
        log_frame.grid_rowconfigure(0, weight=1)
        log_frame.grid_columnconfigure(0, weight=1)
        
        # 로그 텍스트 위젯
        self.log_text = tk.Text(
            log_frame, 
            height=15, 
            wrap=tk.WORD,
            bg="#1a1a1a",
            fg="#00ff88",
            insertbackground="#00ff88",
            selectbackground="#0078d4",
            selectforeground="#ffffff",
            font=('Consolas', 10)
        )
        self.log_text.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        
        # 스크롤바
        scrollbar = ttk.Scrollbar(log_frame, orient=tk.VERTICAL, command=self.log_text.yview)
        scrollbar.grid(row=0, column=1, sticky=(tk.N, tk.S))
        self.log_text.configure(yscrollcommand=scrollbar.set)
        
        # 로그 색상 태그 설정
        self.log_text.tag_configure("info", foreground="#00ff88")
        self.log_text.tag_configure("success", foreground="#4ec9b0")
        self.log_text.tag_configure("warning", foreground="#ffcc00")
        self.log_text.tag_configure("error", foreground="#ff6b6b")
        self.log_text.tag_configure("turbo", foreground="#ff6b00", font=('Consolas', 10, 'bold'))
        
        # 상태바
        self.status_var = tk.StringVar(value="⚡ TURBO READY")
        self.status_bar = ttk.Label(main_frame, textvariable=self.status_var, relief=tk.SUNKEN)
        self.status_bar.grid(row=7, column=0, columnspan=3, sticky=(tk.W, tk.E), pady=(5, 0))
        
        # By 라벨
        by_label = ttk.Label(main_frame, text="By ANYCODER | v3.3 - 대량 크롤링 Edition", foreground='#ff6b00')
        by_label.grid(row=8, column=0, columnspan=3, pady=(5, 0))

    def start_crawling(self):
        keyword = self.search_entry.get().strip()
        if not keyword:
            messagebox.showwarning("입력 오류", "검색어를 입력해주세요.")
            return
        
        try:
            max_count = int(self.max_count_var.get())
            if max_count < 1:
                raise ValueError
        except:
            messagebox.showwarning("입력 오류", "올바른 숫자를 입력해주세요.")
            return
            
        # 대량 크롤링 경고
        if max_count > 100:
            result = messagebox.askyesno(
                "대량 크롤링 경고",
                f"{max_count}개 크롤링을 시작합니다.\n\n"
                "• 메모리 사용량이 증가할 수 있습니다\n"
                "• 헤드리스 모드 사용을 권장합니다\n"
                "• 크래시 발생시 데이터는 저장됩니다\n\n"
                "계속하시겠습니까?"
            )
            if not result:
                return
            
        self.search_button.config(state='disabled', text="⚡ 터보 진행중...")
        self.search_entry.config(state='disabled')
        self.max_count_spinbox.config(state='disabled')
        self.headless_checkbox.config(state='disabled')
        
        # 헤드리스 모드 옵션 전달
        headless_mode = self.headless_var.get()
        self.crawler_thread = CrawlerThread(
            keyword, 
            max_count, 
            self.crawling_finished, 
            self.update_status,
            headless_mode
        )
        self.crawler_thread.root = self.root
        self.crawler_thread.start()
        
    def update_status(self, message):
        self.status_var.set(message)
        
        # 로그창에 메시지 추가
        timestamp = time.strftime("%H:%M:%S")
        log_message = f"[{timestamp}] {message}\n"
        
        # 메시지 타입에 따라 색상 적용
        tag = "info"
        if "⚡" in message or "터보" in message:
            tag = "turbo"
        elif "✅" in message or "완료" in message:
            tag = "success"
        elif "경고" in message or "⚠️" in message:
            tag = "warning"
        elif "오류" in message or "실패" in message or "❌" in message:
            tag = "error"
        
        # 로그 텍스트에 추가
        self.log_text.insert(tk.END, log_message, tag)
        self.log_text.see(tk.END)
        self.root.update_idletasks()
        
    def crawling_finished(self, data):
        self.search_button.config(state='normal', text="⚡ 터보 시작")
        self.search_entry.config(state='normal')
        self.max_count_spinbox.config(state='normal')
        self.headless_checkbox.config(state='normal')
        
        if data:
            self.status_var.set(f"⚡ 터보 크롤링 완료! {len(data)}개 수집!")
            self.save_to_excel(data)
        else:
            messagebox.showinfo("결과 없음", "수집된 데이터가 없습니다.\n검색어를 확인하고 다시 시도해주세요.")
            self.status_var.set("데이터 수집 실패")
            
    def save_to_excel(self, data):
        if not data:
            return
            
        keyword = self.search_entry.get().strip()
        default_filename = f"터보_{keyword}_{time.strftime('%Y%m%d_%H%M%S')}.xlsx"
        
        file_path = filedialog.asksaveasfilename(
            defaultextension=".xlsx",
            filetypes=[("Excel files", "*.xlsx"), ("All files", "*.*")],
            initialfile=default_filename
        )
        
        if file_path:
            try:
                workbook = openpyxl.Workbook()
                sheet = workbook.active
                sheet.title = "터보 크롤링 결과"
                
                # 헤더
                headers = ["번호", "장소명", "도로명 주소", "지번 주소", "전화번호"]
                sheet.append(headers)
                
                # 헤더 스타일
                for cell in sheet[1]:
                    cell.font = openpyxl.styles.Font(bold=True, color="FFFFFF")
                    cell.fill = openpyxl.styles.PatternFill(
                        start_color="FF6B00", 
                        end_color="FF6B00", 
                        fill_type="solid"
                    )
                
                # 데이터 추가
                for idx, item in enumerate(data, 1):
                    sheet.append([idx] + item)
                
                # 열 너비 조정
                for column in sheet.columns:
                    max_length = 0
                    for cell in column:
                        try:
                            if cell.value and len(str(cell.value)) > max_length:
                                max_length = len(str(cell.value))
                        except:
                            pass
                    
                    if hasattr(column[0], 'column_letter'):
                        adjusted_width = min(max_length + 2, 50)
                        sheet.column_dimensions[column[0].column_letter].width = adjusted_width
                    
                workbook.save(file_path)
                
                messagebox.showinfo(
                    "저장 완료",
                    f"⚡ 터보 저장 완료!\n\n"
                    f"파일: {file_path}\n"
                    f"수집 데이터: {len(data)}개"
                )
                self.status_var.set(f"✅ 저장 완료: {file_path}")
                
            except Exception as e:
                messagebox.showerror("저장 실패", f"엑셀 파일 저장 중 오류가 발생했습니다:\n{str(e)}")
                self.status_var.set("저장 실패")


def main():
    root = tk.Tk()
    app = NaverMapCrawlerApp(root)
    
    # 종료 시 크롤링 중지
    def on_closing():
        if app.crawler_thread and app.crawler_thread.is_alive():
            result = messagebox.askquestion(
                "작업 중",
                "크롤링 작업이 아직 진행 중입니다.\n종료하시겠습니까?"
            )
            if result == 'yes':
                app.crawler_thread.stop()
                root.destroy()
        else:
            root.destroy()
    
    root.protocol("WM_DELETE_WINDOW", on_closing)
    root.mainloop()


if __name__ == '__main__':
    main()
