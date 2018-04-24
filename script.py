from selenium import webdriver
from selenium.common.exceptions import NoSuchElementException
from selenium.webdriver.common.keys import Keys
from PIL import Image, ImageChops
import time
import pyautogui
import argparse
import pypandoc
import math


class CiscoConverter:

    def __init__(self, username, password):
        self.driver = webdriver.Chrome(executable_path='chromedriver.exe')
        self.user = username
        self.password = password
        self.section_counter = 0
        self.image_counter = 0
        self.noText = False
        self.section_done = False
        self.first = True
        self.document = open('temp.html', 'w', encoding='utf-8')

    def main(self, filename, start, nr):
        self.driver.set_window_size(1200, 1000)
        self.driver.get('https://www.netacad.com/')
        self.driver.find_element_by_id('headerLoginLink').click()
        self.driver.find_element_by_id('_58_INSTANCE_fm_login').send_keys(self.user)
        self.driver.find_element_by_id('_58_INSTANCE_fm_password').send_keys(self.password)
        self.driver.find_element_by_css_selector('#loginFormModal .btn-primary').click()

        for module in range(start, start+nr):
            print('\nModule ' + str(module))
            self.driver.get('https://static-course-assets.s3.amazonaws.com/IntroNet50ENU/module' +
                            str(module) + '/index.html')
            time.sleep(15)  # Downloading Flash
            self.document.write('<h1>Chapter ' + str(module) + '</h1>')
            while self.driver.find_element_by_id('next-btn').is_enabled():
                self.parse_next_section()

        self.document.close()
        print('\nConverting to docx...')
        pypandoc.convert_file('temp.html', 'docx', outputfile=filename)

    def parse_next_section(self):
        if self.first:
            self.first = False
        else:
            self.driver.find_element_by_id('next-btn').click()
        print('Section ' + str(self.section_counter))
        time.sleep(3)  # Loading Flash
        self.document.write('<h2>' + self.driver.find_element_by_id('title').text + '</h2>')
        self.driver.switch_to.frame(self.driver.find_element_by_id('frame'))  # Content frame
        try:
            self.document.write(self.driver.find_element_by_id('content').get_attribute('innerHTML'))
        except NoSuchElementException:
            print('No text')

        frame = self.driver.find_element_by_id('mainFrame')  # Flash frame
        location = frame.location
        size = frame.size

        while not self.section_done:
            self.screenshot_flash(location, size)

        self.image_counter = 0
        self.section_counter += 1
        self.section_done = False

        self.driver.switch_to.default_content()

    def screenshot_flash(self, location, size):
        time.sleep(1)
        image_name = 'image' + '-' + str(self.section_counter) + '-' + str(self.image_counter)
        self.driver.save_screenshot(image_name + '.png')

        # Cropping values
        left = int(location['x'] + 200)
        top = int(location['y'] + 200)
        right = int(left + size['width'] - 30)
        bottom = int(top + size['height'] - 170)

        if self.image_counter > 0:
            prev_image = Image.open('image' + '-' + str(self.section_counter) + '-' + str(self.image_counter - 1) + '.png')
            current_image = Image.open(image_name + '.png')
            print(self.rmsdiff_2011(prev_image, current_image))
        if (self.image_counter == 0 or self.rmsdiff_2011(prev_image, current_image) > 573) and self.image_counter < 5:
            current_image = Image.open(image_name + '.png')
            current_image = current_image.crop((left, top, right, bottom))
            current_image.save(image_name + '_cropped.png')
            self.document.write('<img src="' + image_name + '_cropped.png" />')
            pyautogui.moveTo(1120, 380 + self.image_counter * 50)
            pyautogui.click()
            self.image_counter += 1
        else:
            pyautogui.moveTo(1120, 330)
            self.section_done = True

    def rmsdiff_2011(self, im1, im2):
        # Calculate the root-mean-square difference between two images
        diff = ImageChops.difference(im1, im2)
        h = diff.histogram()
        sq = (value * (idx ** 2) for idx, value in enumerate(h))
        sum_of_squares = sum(sq)
        rms = math.sqrt(sum_of_squares / float(im1.size[0] * im1.size[1]))
        return rms


parser = argparse.ArgumentParser(description='Convert Netacad courses to Word documents.')
parser.add_argument('username', type=str, help='Your Netacad username')
parser.add_argument('password', type=str, help='Your Netacad password')
parser.add_argument('filename', type=str, help='Filename of the Word document')
parser.add_argument('-s', '--start', nargs='?', type=int, default=0, help='Chapter/Module to start')
parser.add_argument('-n', '--number', nargs='?', type=int, default=1, help='Number of Chapters/Modules')
args = parser.parse_args()

cc = CiscoConverter(args.username, args.password)
cc.main(args.filename, args.start, args.number)
