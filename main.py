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
    def __init__(self, keyword, max_count, callback, status_callback):
        super().__init__()
        self.keyword = keyword
        self.max_count = max_count
        self.callback = callback
        self.status_callback = status_callback
        self.is_running = True
        self.daemon = True

    def find_next_page_button(self, driver):
        """다음 페이지 버튼을 정확히 찾는 함수"""
        # 여러 선택자 시도
        selectors = ["a.eUTV2", "a._2PoiJ", "a[class*='page']", "button[class*='next']"]
        
        for selector in selectors:
            try:
                buttons = driver.find_elements(By.CSS_SELECTOR, selector)
                
                for button in buttons:
                    # aria-disabled 확인
                    if button.get_attribute("aria-disabled") == "true":
                        continue
                    
                    # SVG path로 다음 버튼인지 확인
                    try:
                        svg_element = button.find_element(By.TAG_NAME, "svg")
                        path_element = svg_element.find_element(By.TAG_NAME, "path")
                        d_attribute = path_element.get_attribute("d")
                        
                        # 다음 페이지 화살표 패턴 확인 (오른쪽 화살표)
                        if d_attribute and any(pattern in d_attribute for pattern in ["M12", "M14", "M10.524"]):
                            # 이전 버튼 패턴 제외
                            if not any(pattern in d_attribute for pattern in ["M8", "M6", "M13.476"]):
                                self.status_callback(f"다음 페이지 버튼 찾음: {selector}")
                                return button
                                
                    except:
                        # SVG가 없으면 텍스트나 다른 방법으로 확인
                        button_text = button.text.strip()
                        if "다음" in button_text or ">" in button_text:
                            return button
            except:
                continue
                
        return None

    def scroll_to_load_all(self, driver, scroll_container):
        """모든 아이템이 로드될 때까지 천천히 스크롤"""
        from selenium.webdriver.common.action_chains import ActionChains
        from selenium.webdriver.common.keys import Keys
        
        # 스크롤 가능한 요소 자동 탐지
        if not scroll_container:
            script = """
            var elements = document.querySelectorAll('div[class*="scroll"], div[id*="scroll"], div[role="main"]');
            for (var i = 0; i < elements.length; i++) {
                if (elements[i].scrollHeight > elements[i].clientHeight) {
                    return elements[i];
                }
            }
            return null;
            """
            scroll_container = driver.execute_script(script)
            
            if scroll_container:
                self.status_callback("JavaScript로 스크롤 컨테이너 찾음")
            else:
                # iframe body를 스크롤 컨테이너로 사용
                scroll_container = driver.find_element(By.TAG_NAME, "body")
                self.status_callback("body를 스크롤 컨테이너로 사용")
        
        # 초기 상태 확인
        initial_items = driver.find_elements(By.CSS_SELECTOR, "a.place_bluelink")
        if not initial_items:
            initial_items = driver.find_elements(By.CSS_SELECTOR, "li.UEzoS")
        
        previous_count = len(initial_items)
        self.status_callback(f"초기 아이템 수: {previous_count}개")
        
        # 전체 스크롤 높이 확인
        total_height = driver.execute_script("return arguments[0].scrollHeight", scroll_container)
        current_position = 0
        scroll_step = 300  # 한 번에 스크롤할 픽셀
        
        stable_count = 0
        max_stable_count = 3  # 연속으로 변화가 없을 때 카운트
        
        self.status_callback("천천히 스크롤하며 모든 아이템을 로드합니다...")
        
        # 메인 스크롤 루프
        while current_position < total_height and self.is_running:
            # 점진적 스크롤
            try:
                driver.execute_script(f"""
                    var container = arguments[0];
                    container.scrollTo({{
                        top: {current_position},
                        behavior: 'smooth'
                    }});
                """, scroll_container)
                
                current_position += scroll_step
                
                # 스크롤 후 충분한 대기 시간
                time.sleep(1.5)
                
                # 10번 스크롤마다 더 긴 대기 시간
                if current_position % (scroll_step * 10) == 0:
                    self.status_callback(f"로딩 대기중... (스크롤 위치: {current_position}/{total_height})")
                    time.sleep(2)
                
                # 새로운 아이템 확인
                current_items = driver.find_elements(By.CSS_SELECTOR, "a.place_bluelink")
                if not current_items:
                    current_items = driver.find_elements(By.CSS_SELECTOR, "li.UEzoS")
                
                current_count = len(current_items)
                
                # 새로운 아이템이 로드되었는지 확인
                if current_count > previous_count:
                    self.status_callback(f"새로운 아이템 {current_count - previous_count}개 발견 (총 {current_count}개)")
                    previous_count = current_count
                    stable_count = 0
                    
                    # 새 아이템이 로드되면 전체 높이 다시 확인
                    total_height = driver.execute_script("return arguments[0].scrollHeight", scroll_container)
                else:
                    stable_count += 1
                    
            except Exception as e:
                self.status_callback(f"스크롤 중 오류: {e}")
                pass
        
        # 추가 스크롤 - 끝까지 5번 더 시도
        self.status_callback("추가 스크롤로 놓친 아이템 확인 중...")
        extra_scroll_count = 0
        max_extra_scrolls = 5
        
        while extra_scroll_count < max_extra_scrolls and self.is_running:
            try:
                # 현재 스크롤 위치 확인
                current_scroll = driver.execute_script("return arguments[0].scrollTop", scroll_container)
                
                # 끝까지 스크롤
                driver.execute_script("""
                    var container = arguments[0];
                    container.scrollTo({
                        top: container.scrollHeight,
                        behavior: 'smooth'
                    });
                """, scroll_container)
                time.sleep(2)
                
                # 새로운 스크롤 위치 확인
                new_scroll = driver.execute_script("return arguments[0].scrollTop", scroll_container)
                
                # 스크롤이 더 이상 안 움직이면
                if current_scroll == new_scroll:
                    self.status_callback(f"추가 스크롤 {extra_scroll_count + 1}회 - 더 이상 스크롤 불가")
                else:
                    self.status_callback(f"추가 스크롤 {extra_scroll_count + 1}회 진행")
                
                # 아이템 수 확인
                current_items = driver.find_elements(By.CSS_SELECTOR, "a.place_bluelink")
                if not current_items:
                    current_items = driver.find_elements(By.CSS_SELECTOR, "li.UEzoS")
                
                current_count = len(current_items)
                
                if current_count > previous_count:
                    self.status_callback(f"추가 스크롤에서 {current_count - previous_count}개 더 발견! (총 {current_count}개)")
                    previous_count = current_count
                    # 새로운 아이템을 발견하면 추가 스크롤 카운트 리셋
                    extra_scroll_count = 0
                else:
                    extra_scroll_count += 1
                    
            except Exception as e:
                self.status_callback(f"추가 스크롤 중 오류: {e}")
                extra_scroll_count += 1
        
        # 스크롤을 맨 위로
        try:
            driver.execute_script("arguments[0].scrollTop = 0", scroll_container)
        except:
            driver.execute_script("window.scrollTo(0, 0);")
        
        time.sleep(1)  # 스크롤 복귀 대기
        
        # 최종 아이템 수 반환
        final_items = driver.find_elements(By.CSS_SELECTOR, "a.place_bluelink")
        if not final_items:
            final_items = driver.find_elements(By.CSS_SELECTOR, "li.UEzoS")
        
        self.status_callback(f"스크롤 완료. 최종 로드된 아이템: {len(final_items)}개")
        return len(final_items)

    def run(self):
        data = []
        collected_count = 0
        self.status_callback(f"'{self.keyword}' 검색을 시작합니다... (최대 {self.max_count}개 수집)")
        driver = None

        try:
            options = webdriver.ChromeOptions()
            # options.add_argument("--headless") # Keep this commented for debugging
            options.add_argument("--log-level=3")
            options.add_argument("--disable-blink-features=AutomationControlled")
            options.add_experimental_option("excludeSwitches", ["enable-automation"])
            options.add_experimental_option('useAutomationExtension', False)
            driver = webdriver.Chrome(options=options)

            encoded_keyword = quote(self.keyword)
            search_url = f"https://map.naver.com/p/search/{encoded_keyword}"
            driver.get(search_url)
            time.sleep(3)

            WebDriverWait(driver, 20).until(EC.frame_to_be_available_and_switch_to_it((By.ID, "searchIframe")))
            time.sleep(2)

            current_page = 1
            max_pages_to_check = 20 # Limit to prevent infinite loops on pagination issues

            while self.is_running and collected_count < self.max_count and current_page <= max_pages_to_check:
                self.status_callback(f"\n===== {current_page} 페이지 크롤링 시작 =====")

                # Scroll to load all elements on the current page
                scroll_container = None
                scroll_selectors = [
                    "div#_list_scroll_container",  # 가장 정확한 선택자
                    "div._2ky45",  # 대안 1
                    "div[class*='scroll_area']",  # 대안 2
                    "div[role='main']",  # 대안 3
                    "div.Ryr1F",  # 대안 4
                    "#searchIframe"  # iframe 자체
                ]
                
                for selector in scroll_selectors:
                    try:
                        if selector == "#searchIframe":
                            # iframe 자체를 스크롤하는 경우
                            scroll_container = driver.find_element(By.TAG_NAME, "body")
                        else:
                            scroll_container = driver.find_element(By.CSS_SELECTOR, selector)
                        
                        if scroll_container:
                            # 스크롤 가능한지 확인
                            scroll_height = driver.execute_script("return arguments[0].scrollHeight", scroll_container)
                            client_height = driver.execute_script("return arguments[0].clientHeight", scroll_container)
                            
                            if scroll_height > client_height:
                                self.status_callback(f"스크롤 컨테이너 찾음: {selector}")
                                break
                    except:
                        continue

                if not scroll_container:
                    self.status_callback("스크롤 컨테이너를 찾을 수 없음 (body 스크롤 사용)")
                    scroll_container = driver.find_element(By.TAG_NAME, "body")

                # 모든 아이템 로드
                self.scroll_to_load_all(driver, scroll_container)

                # Get all place elements on the current page after scrolling
                all_place_elements = []
                # 여러 선택자 시도
                place_selectors = [
                    "a.place_bluelink",  # 최신 선택자 - 클릭 가능한 링크
                    "li.UEzoS.rTjJo",  # 리스트 아이템
                    "li._1EKsQ._12tNp",  # 대안 1
                    "li.UEzoS",  # 기존 선택자
                    "a[class*='place_bluelink']"  # 부분 매칭
                ]
                
                for selector in place_selectors:
                    try:
                        all_place_elements = WebDriverWait(driver, 5).until(
                            EC.presence_of_all_elements_located((By.CSS_SELECTOR, selector))
                        )
                        if all_place_elements:
                            self.status_callback(f"장소 요소 찾음: {selector} ({len(all_place_elements)}개)")
                            break
                    except TimeoutException:
                        continue
                    except Exception as e:
                        continue
                
                if not all_place_elements:
                    self.status_callback(f"{current_page} 페이지에서 장소를 찾을 수 없습니다.")
                    # If no elements found, try to move to next page
                    pass

                if not all_place_elements:
                    self.status_callback(f"{current_page} 페이지에 크롤링할 장소가 없습니다. 다음 페이지로 이동 시도.")
                    pass_to_next_page = True
                else:
                    pass_to_next_page = False
                    self.status_callback(f"{current_page} 페이지에서 총 {len(all_place_elements)}개의 장소를 찾았습니다.")

                    # Iterate through each place element and crawl its details
                    for i, element in enumerate(all_place_elements):
                        if not self.is_running or collected_count >= self.max_count:
                            break

                        place_name_for_log = "알 수 없는 장소"
                        try:
                            # a.place_bluelink를 직접 클릭하는 경우
                            if element.tag_name == 'a':
                                element_to_click = element
                                try:
                                    name_span = element.find_element(By.CSS_SELECTOR, "span.YwYLL")
                                    place_name_for_log = name_span.text.strip()
                                except:
                                    pass
                            else:
                                # li 요소인 경우 내부의 a.place_bluelink 찾기
                                element_to_click = element.find_element(By.CSS_SELECTOR, "a.place_bluelink")
                                try:
                                    name_span = element_to_click.find_element(By.CSS_SELECTOR, "span.YwYLL")
                                    place_name_for_log = name_span.text.strip()
                                except:
                                    pass

                            driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", element_to_click)
                            time.sleep(0.5)
                            driver.execute_script("arguments[0].click();", element_to_click)
                            time.sleep(2)

                            driver.switch_to.default_content()
                            WebDriverWait(driver, 10).until(EC.frame_to_be_available_and_switch_to_it((By.ID, "entryIframe")))
                            time.sleep(1)

                            name, road_address, jibun_address, phone = "정보 없음", "정보 없음", "정보 없음", "정보 없음"

                            # 장소명 - 여러 선택자 시도
                            name_selectors = [".YwYLL", "._3Apjo", ".GHAhO", "h2", "[class*='title']"]
                            for selector in name_selectors:
                                try:
                                    name_elem = WebDriverWait(driver, 3).until(
                                        EC.presence_of_element_located((By.CSS_SELECTOR, selector))
                                    )
                                    name = name_elem.text.strip().replace('복사', '').strip()
                                    if name and name != "정보 없음":
                                        break
                                except:
                                    continue

                            # 주소 - PkgBl 버튼 클릭해서 지번주소 가져오기
                            try:
                                # 먼저 도로명 주소 가져오기
                                road_address_elem = driver.find_element(By.CSS_SELECTOR, ".LDgIH")
                                road_address = road_address_elem.text.strip().replace('복사', '').strip()
                                
                                # PkgBl 버튼 클릭해서 지번주소 메뉴 열기
                                try:
                                    address_button = driver.find_element(By.CSS_SELECTOR, "a.PkgBl")
                                    driver.execute_script("arguments[0].click();", address_button)
                                    time.sleep(0.5)
                                    
                                    # nQ7Lh 클래스에서 지번주소 찾기
                                    address_items = driver.find_elements(By.CSS_SELECTOR, ".nQ7Lh")
                                    for item in address_items:
                                        item_text = item.text.strip()
                                        if "지번" in item_text:
                                            # "지번"과 "복사" 문구 제거
                                            jibun_address = item_text.replace("지번", "").replace("복사", "").strip()
                                            break
                                except:
                                    # 버튼 클릭 실패시 기존 방식으로 시도
                                    address_selectors = [".LDgIH", "._2yqUQ", "span[class*='addr']"]
                                    for selector in address_selectors:
                                        try:
                                            address_elems = driver.find_elements(By.CSS_SELECTOR, selector)
                                            if len(address_elems) > 1:
                                                jibun_address = address_elems[1].text.strip().replace('복사', '').strip()
                                                break
                                        except:
                                            continue
                            except Exception as e:
                                self.status_callback(f"주소 가져오기 오류: {e}")

                            # 전화번호 - 여러 선택자 시도
                            phone_selectors = [".xlx7Q", "._3ZA58 span", ".dry01", ".J7eF_", "span[class*='phone']", "[class*='tel']"]
                            phone_found = False
                            
                            for selector in phone_selectors:
                                try:
                                    phone_elems = driver.find_elements(By.CSS_SELECTOR, selector)
                                    if phone_elems:
                                        phone_text = phone_elems[0].text.strip()
                                        # "휴대전화번호"와 "복사" 문구 제거
                                        phone_text = phone_text.replace("휴대전화번호", "").replace("복사", "").strip()
                                        # 전화번호 형식 확인 (숫자와 하이픈 포함)
                                        if any(char.isdigit() for char in phone_text):
                                            phone = phone_text
                                            phone_found = True
                                            break
                                except:
                                    continue
                            
                            # 전화번호를 못 찾았으면 "전화번호 보기" 버튼 클릭
                            if not phone_found:
                                try:
                                    # 전화번호 보기 버튼 찾기
                                    phone_button = driver.find_element(By.CSS_SELECTOR, "a.BfF3H")
                                    if "전화번호 보기" in phone_button.text:
                                        self.status_callback("전화번호 보기 버튼 클릭")
                                        driver.execute_script("arguments[0].click();", phone_button)
                                        time.sleep(0.5)
                                        
                                        # 클릭 후 전화번호 찾기
                                        try:
                                            phone_elem = driver.find_element(By.CSS_SELECTOR, ".J7eF_")
                                            phone_text = phone_elem.text.strip()
                                            # "휴대전화번호"와 "복사" 문구 제거
                                            phone = phone_text.replace("휴대전화번호", "").replace("복사", "").strip()
                                            if not any(char.isdigit() for char in phone):
                                                phone = "정보 없음"
                                        except:
                                            # 다른 선택자들도 시도
                                            for selector in phone_selectors:
                                                try:
                                                    phone_elems = driver.find_elements(By.CSS_SELECTOR, selector)
                                                    if phone_elems:
                                                        phone_text = phone_elems[0].text.strip()
                                                        # "휴대전화번호"와 "복사" 문구 제거
                                                        phone_text = phone_text.replace("휴대전화번호", "").replace("복사", "").strip()
                                                        if any(char.isdigit() for char in phone_text):
                                                            phone = phone_text
                                                            break
                                                except:
                                                    continue
                                except Exception as e:
                                    self.status_callback(f"전화번호 보기 버튼 처리 중 오류: {e}")

                            if name != "정보 없음":
                                data.append([name, road_address, jibun_address, phone])
                                collected_count += 1
                                self.status_callback(f"({collected_count}/{self.max_count}) {name} 정보 수집 완료")
                            else:
                                self.status_callback(f"장소 이름(YwYLL)을 찾을 수 없어 정보 수집 건너뜀.")

                        except StaleElementReferenceException:
                            self.status_callback(f"StaleElementReferenceException 발생. '{place_name_for_log}' 건너뜀")
                            continue
                        except Exception as e:
                            self.status_callback(f"장소 '{place_name_for_log}' 처리 중 오류: {str(e)}")
                            continue
                        finally:
                            # Always switch back to searchIframe for the next iteration
                            driver.switch_to.default_content()
                            WebDriverWait(driver, 10).until(EC.frame_to_be_available_and_switch_to_it((By.ID, "searchIframe")))
                            time.sleep(0.5)

                if collected_count >= self.max_count:
                    self.status_callback(f"목표 개수({self.max_count}개)를 모두 수집하여 크롤링을 종료합니다.")
                    break

                # Move to the next page - 개선된 로직
                try:
                    driver.switch_to.default_content()
                    WebDriverWait(driver, 10).until(EC.frame_to_be_available_and_switch_to_it((By.ID, "searchIframe")))
                    
                    # 페이지네이션 컨테이너로 스크롤
                    try:
                        pagination_container = WebDriverWait(driver, 5).until(
                            EC.presence_of_element_located((By.CSS_SELECTOR, "div.zRM9F"))
                        )
                        driver.execute_script("arguments[0].scrollIntoView(true);", pagination_container)
                        time.sleep(0.5)
                    except:
                        self.status_callback("페이지네이션 컨테이너를 찾을 수 없음")
                    
                    # 다음 페이지 버튼 찾기
                    next_button = self.find_next_page_button(driver)
                    
                    if next_button:
                        driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", next_button)
                        time.sleep(0.3)
                        driver.execute_script("arguments[0].click();", next_button)
                        self.status_callback(f"\n{current_page + 1} 페이지로 이동합니다.")
                        current_page += 1
                        time.sleep(3)
                        
                        # 새 페이지 로드 확인
                        try:
                            WebDriverWait(driver, 5).until(
                                lambda d: len(d.find_elements(By.CSS_SELECTOR, "a.place_bluelink")) > 0
                            )
                        except:
                            self.status_callback("새 페이지 로드 대기 중...")
                    else:
                        # 페이지 번호로 직접 이동 시도
                        try:
                            current_page_elem = driver.find_element(By.CSS_SELECTOR, "a[aria-current='true']")
                            current_num = int(current_page_elem.text.strip())
                            next_num = current_num + 1
                            
                            next_page_buttons = driver.find_elements(By.XPATH, f"//a[text()='{next_num}']")
                            if next_page_buttons:
                                driver.execute_script("arguments[0].click();", next_page_buttons[0])
                                self.status_callback(f"\n{next_num} 페이지로 이동합니다.")
                                current_page = next_num
                                time.sleep(3)
                                
                                if status_callback:
                                    status_callback(f"\n{current_page} 페이지로 이동합니다.")
                            else:
                                self.status_callback("마지막 페이지에 도달하여 크롤링을 종료합니다.")
                                break
                        except:
                            self.status_callback("다음 페이지를 찾을 수 없습니다. 마지막 페이지일 수 있습니다.")
                            break
                            
                except (NoSuchElementException, TimeoutException) as e:
                    self.status_callback(f"페이지네이션 요소를 찾을 수 없습니다: {e}")
                    self.status_callback("마지막 페이지에 도달하여 크롤링을 종료합니다.")
                    break
                except Exception as e:
                    self.status_callback(f"다음 페이지 이동 중 오류 발생: {e}")
                    # 오류 발생시 스크린샷 저장 (디버깅용)
                    try:
                        driver.save_screenshot(f"error_page_{current_page}.png")
                        self.status_callback(f"오류 스크린샷 저장: error_page_{current_page}.png")
                    except:
                        pass
                    break

            self.status_callback(f"\n크롤링 완료. 총 {collected_count}개의 정보를 수집했습니다.")

        except Exception as e:
            self.status_callback(f"크롤링 프로세스 중 심각한 오류 발생: {str(e)}")
        finally:
            if driver:
                driver.quit()
            self.root.after(0, self.callback, data)

    def stop(self):
        self.is_running = False


class NaverMapCrawlerApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Naver Map Crawler v2.0.1 - Final")
        self.root.geometry("700x300")
        
        # 스타일 설정
        style = ttk.Style()
        style.configure('Title.TLabel', font=('Arial', 18, 'bold'))
        
        self.crawler_thread = None
        self.setup_ui()
        
    def setup_ui(self):
        # 메인 프레임
        main_frame = ttk.Frame(self.root, padding="20")
        main_frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        
        # 타이틀
        title_label = ttk.Label(main_frame, text="네이버 지도 크롤러 (최종버전)", style='Title.TLabel')
        title_label.grid(row=0, column=0, columnspan=3, pady=(0, 20))
        
        # 검색 영역
        search_frame = ttk.Frame(main_frame)
        search_frame.grid(row=1, column=0, columnspan=3, pady=10)
        
        ttk.Label(search_frame, text="검색어:").grid(row=0, column=0, padx=5)
        self.search_entry = ttk.Entry(search_frame, width=30)
        self.search_entry.grid(row=0, column=1, padx=5)
        self.search_entry.bind('<Return>', lambda e: self.start_crawling())
        
        ttk.Label(search_frame, text="최대 갯수:").grid(row=0, column=2, padx=5)
        self.max_count_var = tk.StringVar(value="100")
        self.max_count_spinbox = ttk.Spinbox(search_frame, from_=1, to=500, textvariable=self.max_count_var, width=10)
        self.max_count_spinbox.grid(row=0, column=3, padx=5)
        
        self.search_button = ttk.Button(search_frame, text="검색 시작", command=self.start_crawling)
        self.search_button.grid(row=0, column=4, padx=10)
        
        # 안내 메시지
        info_text = "• 검색 결과를 엑셀 파일로 저장합니다.\n• 수집 항목: 장소명, 도로명 주소, 지번 주소, 전화번호\n• 최대 500개까지 수집 가능 (여러 페이지 자동 크롤링)\n• 전화번호 보기 버튼도 자동 처리됩니다."
        info_label = ttk.Label(main_frame, text=info_text, foreground='gray')
        info_label.grid(row=2, column=0, columnspan=3, pady=20)
        
        # 상태바
        self.status_var = tk.StringVar(value="준비 완료")
        self.status_bar = ttk.Label(main_frame, textvariable=self.status_var, relief=tk.SUNKEN)
        self.status_bar.grid(row=3, column=0, columnspan=3, sticky=(tk.W, tk.E), pady=(20, 0))
        
        # By 라벨
        by_label = ttk.Label(main_frame, text="By ANYCODER | v2.0.1 Final", foreground='gray')
        by_label.grid(row=4, column=0, columnspan=3, pady=(10, 0))
        
    def start_crawling(self):
        keyword = self.search_entry.get().strip()
        if not keyword:
            messagebox.showwarning("입력 오류", "검색어를 입력해주세요.")
            return
        
        try:
            max_count = int(self.max_count_var.get())
            if max_count < 1 or max_count > 500:
                raise ValueError
        except:
            messagebox.showwarning("입력 오류", "최대 갯수는 1~500 사이의 숫자여야 합니다.")
            return
            
        self.search_button.config(state='disabled', text="크롤링 중...")
        self.search_entry.config(state='disabled')
        self.max_count_spinbox.config(state='disabled')
        
        self.crawler_thread = CrawlerThread(keyword, max_count, self.crawling_finished, self.update_status)
        self.crawler_thread.root = self.root  # root 참조 추가
        self.crawler_thread.start()
        
    def update_status(self, message):
        self.status_var.set(message)
        self.root.update_idletasks()
        
    def crawling_finished(self, data):
        self.search_button.config(state='normal', text="검색 시작")
        self.search_entry.config(state='normal')
        self.max_count_spinbox.config(state='normal')
        
        if data:
            self.status_var.set(f"크롤링 완료. {len(data)}개의 정보를 수집했습니다.")
            self.save_to_excel(data)
        else:
            messagebox.showinfo("결과 없음", "수집된 데이터가 없습니다.\n검색어를 확인하고 다시 시도해주세요.")
            self.status_var.set("데이터 수집 실패")
            
    def save_to_excel(self, data):
        if not data:
            return
            
        keyword = self.search_entry.get().strip()
        default_filename = f"네이버지도_{keyword}_{time.strftime('%Y%m%d_%H%M%S')}.xlsx"
        
        file_path = filedialog.asksaveasfilename(
            defaultextension=".xlsx",
            filetypes=[("Excel files", "*.xlsx"), ("All files", "*.*")],
            initialfile=default_filename
        )
        
        if file_path:
            try:
                workbook = openpyxl.Workbook()
                sheet = workbook.active
                sheet.title = "네이버 지도 크롤링 결과"
                
                # 헤더
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
                    f"'{file_path}'에 성공적으로 저장되었습니다.\n\n"
                    f"수집된 데이터: {len(data)}개"
                )
                self.status_var.set(f"저장 완료: {file_path}")
                
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
