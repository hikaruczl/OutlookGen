import socket as dsocket
from contextlib import suppress
from ctypes import windll
from dataclasses import dataclass
from json import load
from os import getcwd, listdir, chdir, remove, mkdir, path, name
from random import choice, randrange
from random import randint
from shutil import copytree, rmtree, copyfile
from ssl import create_default_context
from sys import exit
from time import sleep, time
from typing import Any
from urllib.parse import urlparse
from zipfile import ZipFile

from colorama import Fore
from requests import get
from playwright.sync_api import sync_playwright, Page, Browser
from tqdm import tqdm
from twocaptcha import TwoCaptcha

from Utils import Utils, Timer
from anycaptcha import AnycaptchaClient, FunCaptchaProxylessTask

# Some Value
eGenerated = 0
solvedCaptcha = 0


class AutoUpdater:
    def __init__(self, version):
        self.version = version
        self.latest = self.get_latest()
        self.this = getcwd()
        self.file = "temp/latest.zip"
        self.folder = f"temp/latest_{randrange(1_000_000, 999_999_999)}"

    @dataclass
    class latest_data:
        version: str
        zip_url: str

    def get_latest(self):
        rjson = get("https://api.github.com/repos/MatrixTM/OutlookGen/tags").json()
        return self.latest_data(version=rjson[0]["name"], zip_url=get(rjson[0]["zipball_url"]).url)

    @staticmethod
    def download(host, dPath, filename):
        with dsocket.socket(dsocket.AF_INET, dsocket.SOCK_STREAM) as sock:
            context = create_default_context()
            with context.wrap_socket(sock, server_hostname="api.github.com") as wrapped_socket:
                wrapped_socket.connect((dsocket.gethostbyname(host), 443))
                wrapped_socket.send(
                    f"GET {dPath} HTTP/1.1\r\nHost:{host}\r\nAccept: text/html,application/xhtml+xml,application/xml;q=0.9,file/avif,file/webp,file/apng,*/*;q=0.8,application/signed-exchange;v=b3;q=0.9\r\n\r\n".encode())

                resp = b""
                while resp[-4:-1] != b"\r\n\r":
                    resp += wrapped_socket.recv(1)
                else:
                    resp = resp.decode()
                    content_length = int(
                        "".join([tag.split(" ")[1] for tag in resp.split("\r\n") if "content-length" in tag.lower()]))
                    _file = b""
                    while content_length > 0:
                        data = wrapped_socket.recv(2048)
                        if not data:
                            print("EOF")
                            break
                        _file += data
                        content_length -= len(data)
                    with open(filename, "wb") as file:
                        file.write(_file)

    def update(self):
        if not self.version == self.latest.version:
            rmtree("temp") if path.exists("temp") else ""
            mkdir("temp")
            print("Updating Script...")
            parsed = urlparse(self.latest.zip_url)
            self.download(parsed.hostname, parsed.path, self.file)
            ZipFile(self.file).extractall(self.folder)
            print(path.exists(self.folder))
            print(path.exists(listdir(self.folder)[0]))
            chdir("{}/{}".format(self.folder, listdir(self.folder)[0]))
            for files in listdir():
                if path.isdir(files):
                    with suppress(FileNotFoundError):
                        rmtree("{}/{}".format(self.this, files))
                    copytree(files, "{}/{}".format(self.this, files))
                else:
                    with suppress(FileNotFoundError):
                        remove("{}/{}".format(self.this, files))
                    copyfile(files, "{}/{}".format(self.this, files))
            rmtree("../../../temp")
            exit("Run Script Again!")
            return
        print("Script is up to date!")


class eGen:
    def __init__(self):
        self.version = "v1.2.4"
        AutoUpdater(self.version).update()
        self.Utils = Utils()  # Utils Module
        self.config: Any = load(open('config.json'))  # Config File
        self.checkConfig()  # Check Config File
        self.Timer = Timer()  # Timer

        self.playwright = None
        self.browser = None
        self.first_name = None  # Generate First Name
        self.last_name = None  # Generate Last Name
        self.password = None  # Generate Password
        self.email = None  # Generate Email

        # Values About Captcha
        self.providers = self.config['Captcha']['providers']
        self.api_key = self.config["Captcha"]["api_key"]
        self.site_key = self.config["Captcha"]["site_key"]

        # Other
        self.proxies = [i.strip() for i in open(self.config['Common']['ProxyFile']).readlines()]  # Get Proxies

        # Browser launch arguments (converted from Chrome args)
        self.browser_args = []
        for arg in tqdm(self.config["DriverArguments"], desc='Loading Arguments',
                        bar_format='{desc} | {l_bar}{bar:15} | {percentage:3.0f}%'):
            if arg not in ['--disable-blink-features=AutomationControlled']:  # Playwright handles this differently
                self.browser_args.append(arg)
            sleep(0.2)

    def solver(self, site_url, page):
        # Solve Captcha Function
        global solvedCaptcha
        # TwoCaptcha
        if self.providers == 'twocaptcha':
            try:
                return TwoCaptcha(self.api_key).funcaptcha(sitekey=self.site_key, url=site_url)['code']
            except Exception as exp:
                self.print(exp)

            # AnyCaptcha
        elif self.providers == 'anycaptcha':
            client = AnycaptchaClient(self.api_key)
            task = FunCaptchaProxylessTask(site_url, self.site_key)
            job = client.createTask(task, typecaptcha="funcaptcha")
            self.print("Solving funcaptcha")
            job.join()
            result = job.get_solution_response()
            if result.find("ERROR") != -1:
                self.print(result)
                page.close()
            else:
                solvedCaptcha += 1
                return result

    def check_proxy(self, proxy):
        with suppress(Exception):
            get("https://outlook.live.com", proxies={
                "http": "http://{}".format(proxy),
                "https": "http://{}".format(proxy)
            }, timeout=self.config["Common"]["ProxyCheckTimeout"] or 5)
            return True
        return False

    def fElement(self, page: Page, selector: str, timeout: float = 30000):
        # Custom find Element Function using Playwright
        try:
            return page.locator(selector)
        except Exception as e:
            self.print(f'Failed to find element: {selector}')
            page.close()
            return None

    def get_balance(self):
        # Check provider Balance Function
        if self.providers == 'twocaptcha':
            return TwoCaptcha(self.api_key).balance()
        elif self.providers == 'anycaptcha':
            return AnycaptchaClient(self.api_key).getBalance()

    def update(self):
        # Update Title Function
        global eGenerated, solvedCaptcha
        title = f'Email Generated: {eGenerated} | Solved Captcha: {solvedCaptcha} | Balance: {self.get_balance()}'
        windll.kernel32.SetConsoleTitleW(title) if name == 'nt' else print(f'\33]0;{title}\a', end='',
                                                                           flush=True)

    def generate_info(self):
        # Generate Information Function
        self.email = self.Utils.eGen()
        self.password = self.Utils.makeString(self.config["EmailInfo"]["PasswordLength"])  # Generate Password
        self.first_name = self.Utils.makeString(self.config["EmailInfo"]["FirstNameLength"])  # Generate First Name
        self.last_name = self.Utils.makeString(self.config["EmailInfo"]["LastNameLength"])  # Generate Last Name

    def checkConfig(self):
        # Check Config Function
        captcha_sec = self.config['Captcha']
        if captcha_sec['api_key'] == "" or captcha_sec['providers'] == "anycaptcha/twocaptcha" or \
                self.config['EmailInfo']['Domain'] == "@hotmail.com/@outlook.com":
            self.print('Please Fix Config!')
            exit()

    def print(self, text: object, end: str = "\n"):
        # Print With Prefix Function
        print(self.Utils.replace(f"{self.config['Common']['Prefix']}&f{text}",
                                 {
                                     '&a': Fore.LIGHTGREEN_EX,
                                     '&4': Fore.RED,
                                     '&2': Fore.GREEN,
                                     '&b': Fore.LIGHTCYAN_EX,
                                     '&c': Fore.LIGHTRED_EX,
                                     '&6': Fore.LIGHTYELLOW_EX,
                                     '&f': Fore.RESET,
                                     '&e': Fore.LIGHTYELLOW_EX,
                                     '&3': Fore.CYAN,
                                     '&1': Fore.BLUE,
                                     '&9': Fore.LIGHTBLUE_EX,
                                     '&5': Fore.MAGENTA,
                                     '&d': Fore.LIGHTMAGENTA_EX,
                                     '&8': Fore.LIGHTBLACK_EX,
                                     '&0': Fore.BLACK}), end=end)

    def CreateEmail(self, page: Page):
        # Create Email Function
        try:
            global eGenerated, solvedCaptcha
            self.update()
            self.Timer.start(time()) if self.config["Common"]['Timer'] else ''

            page.goto("https://outlook.live.com/owa/?nlp=1&signup=1")
            page.wait_for_load_state("networkidle")
            assert 'Create' in page.title()

            if self.config['EmailInfo']['Domain'] == "@hotmail.com":
                domain = page.locator('#LiveDomainBoxList')
                domain.select_option('hotmail.com')
                sleep(0.1)

            emailInput = page.locator('#MemberName')
            emailInput.fill(self.email)
            self.print(f"email: {self.email}{self.config['EmailInfo']['Domain']}")
            page.locator('#iSignupAction').click()

            with suppress(Exception):
                error_text = page.locator('#MemberNameError').text_content(timeout=2000)
                if error_text:
                    self.print(error_text)
                    self.print("email is already taken")
                    page.close()
                    return

            passwordinput = page.locator('#PasswordInput')
            passwordinput.fill(self.password)
            self.print("Password: %s" % self.password)
            page.locator('#iSignupAction').click()

            first = page.locator("#FirstName")
            first.fill(self.first_name)
            sleep(.3)
            last = page.locator("#LastName")
            last.fill(self.last_name)
            page.locator('#iSignupAction').click()

            dropdown = page.locator("#Country")
            dropdown.select_option(label='Turkey')

            birthMonth = page.locator("#BirthMonth")
            birthMonth.select_option(str(randint(1, 12)))

            birthDay = page.locator("#BirthDay")
            birthDay.select_option(str(randint(1, 28)))

            birthYear = page.locator("#BirthYear")
            birthYear.fill(str(randint(self.config['EmailInfo']['minBirthDate'], self.config['EmailInfo']['maxBirthDate'])))

            page.locator('#iSignupAction').click()

            # Handle iframe for captcha
            frame = page.frame_locator('#enforcementFrame')
            token = self.solver(page.url, page)
            sleep(0.5)
            page.evaluate(
                f'parent.postMessage(JSON.stringify({{eventId:"challenge-complete",payload:{{sessionToken:"{token}"}}}}),"*")')
            self.print("&aCaptcha Solved")
            self.update()

            page.locator('#idBtn_Back').click()
            self.print(f'Email Created in {str(self.Timer.timer(time())).split(".")[0]}s') if \
                self.config["Common"]['Timer'] else self.print('Email Created')
            eGenerated += 1
            self.Utils.logger(self.email + self.config['EmailInfo']['Domain'], self.password)
            self.update()
            page.close()
        except Exception as e:
            if e == KeyboardInterrupt:
                page.close()
                exit(0)
            self.print("&4Something is wrong | %s" % str(e).split("\n")[0].strip())
        finally:
            page.close()

    def get_valid_proxy(self):
        """Get a valid proxy from the proxy list"""
        while self.proxies:
            proxy = choice(self.proxies)
            if self.check_proxy(proxy):
                self.print(f"&aUsing proxy: &f{proxy}")
                return proxy
            else:
                self.print("&c%s &f| &4Invalid Proxy&f" % proxy)
                self.proxies.remove(proxy)
        return None

    def create_single_account(self):
        """Create a single email account (one complete registration flow)"""
        # Generate account information
        self.generate_info()

        # Get a valid proxy
        proxy = self.get_valid_proxy()
        if not proxy:
            self.print("&4No valid proxy available!")
            return False

        # Create the email account using Playwright
        try:
            with sync_playwright() as p:
                # Parse proxy format: ip:port or ip:port:username:password
                proxy_parts = proxy.split(':')
                proxy_config = {
                    "server": f"http://{proxy_parts[0]}:{proxy_parts[1]}"
                }
                if len(proxy_parts) == 4:
                    proxy_config["username"] = proxy_parts[2]
                    proxy_config["password"] = proxy_parts[3]

                # Launch browser with proxy
                browser = p.chromium.launch(
                    headless='--headless' in self.browser_args,
                    args=self.browser_args,
                    proxy=proxy_config
                )

                # Create context and page
                context = browser.new_context()
                page = context.new_page()

                # Create the email
                self.CreateEmail(page=page)

                # Cleanup
                context.close()
                browser.close()
            return True
        except Exception as e:
            self.print(f"&4Failed to create account: {str(e)}")
            return False

    def run(self, count=None):
        """
        Run the email generation script

        Args:
            count (int, optional): Number of accounts to create.
                                   If None, runs indefinitely until proxies run out.
        """
        self.print('&bCoded with &c<3&b by MatrixTeam')

        if count is None:
            self.print('&eRunning in unlimited mode (Ctrl+C to stop)')
        else:
            self.print(f'&eWill create &a{count}&e account(s)')

        created = 0
        attempts = 0

        with suppress(IndexError):
            while True:
                # Check if we've reached the target count
                if count is not None and created >= count:
                    self.print(f'&aCompleted! Created {created} account(s)')
                    break

                # Check if we have proxies available
                if not self.proxies:
                    self.print("&4No Proxy Available, Exiting!")
                    break

                attempts += 1
                self.print(f'\n&b--- Attempt {attempts} (Created: {created}/{count if count else "unlimited"}) ---')

                # Try to create an account
                if self.create_single_account():
                    created += 1

        self.print(f'\n&b=== Summary ===')
        self.print(f'&aSuccessfully created: {created}')
        self.print(f'&eTotal attempts: {attempts}')



if __name__ == '__main__':
    from sys import argv

    # Check if user provided a count argument
    if len(argv) > 1:
        try:
            count = int(argv[1])
            if count <= 0:
                print("Error: Count must be a positive number")
                exit(1)
            eGen().run(count=count)
        except ValueError:
            print(f"Error: Invalid count '{argv[1]}'. Please provide a positive integer.")
            print("Usage: python main.py [count]")
            print("  count: Number of accounts to create (optional, default: unlimited)")
            exit(1)
    else:
        # No argument provided, run unlimited mode
        eGen().run()
