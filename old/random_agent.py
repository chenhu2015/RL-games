
import mss, mss.tools
import numpy as np
import os
from PIL import ImageOps, Image
import time
import win32api, win32con, win32gui, win32com.client
import matplotlib.pyplot as plt

shell = win32com.client.Dispatch("WScript.Shell")
shell.SendKeys('%')

def main():
    # Create folder to save images
    try: 
        os.mkdir("images") 
    except OSError:
        pass
    
    action_space = ActionSpace()
    window_getter = WindowGetter()
    score = Score()
    
    start_time = time.time() 
    done = False
    
    while not done:
        win32gui.EnumWindows(window_getter.callback, None)
        
        sct_img, array = window_getter.screenshot() 
        image = Image.new("RGB", sct_img.size)
        pixels = zip(sct_img.raw[2::4], sct_img.raw[1::4], sct_img.raw[0::4])
        image.putdata(list(pixels))
        
        if score.is_image_of_game_over(array) and score.current_score > 0:
            score.reset_scoring_function()
            print(f"Score: {score.current_score}")
        
        action_space.get_random_action()  
        
        score_image = image.crop(score.score_bounding_box)
        if score.score_changed(np.array(score_image)):
            print(f"Score: {score.current_score}")
        
        # Currently runs for 60 seconds
        if time.time() > start_time + 60:
            done = True
            
            
class Score:
    
    def __init__(self):
        self.coords_of_pixel_in_G_of_GameOver = (110, 80)
        self.coords_of_pixel_under_G_of_GameOver = (115, 80)
        self.coords_of_pixel_in_R = (110, 168)
        self.score_bounding_box = (108, 36, 148, 54) 
        self.num_previous_score_images = 5
        self.last_score_images_white = [None] * self.num_previous_score_images
        self.last_score_images_black = [None] * self.num_previous_score_images
        self.last_score_image_index = 0
        self.last_score_images_stacked_white = None
        self.last_score_images_stacked_black = None
        self.require_change_in_a_row = 5
        self.current_change_in_a_row = 0
        self.current_score = 0
        
    def is_image_of_game_over(self, im):
        a = self.coords_of_pixel_in_G_of_GameOver
        b = self.coords_of_pixel_under_G_of_GameOver
        c = self.coords_of_pixel_in_R
        # plt.imshow(im)
        # plt.plot(a[0], a[1], 'rx')
        # plt.show()
        # input("continue: ")
        return (im[a[0]][a[1]][0] == 255 and im[b[0]][b[1]][0] == 0 and im[c[0]][c[1]][0] == 255)

    # Used to reset all the global variables to their initial state
    def reset_scoring_function(self):
      self.last_score_images_white = [None] * self.num_previous_score_images
      self.last_score_images_black = [None] * self.num_previous_score_images
      self.last_score_image_index = 0
      self.last_score_images_stacked_white = None
      self.last_score_images_stacked_black = None
      self.current_change_in_a_row = 0
      self.current_score = 0

    # Used to check if the score has changed. Need a few frames in a row to have been different before it is registered
    def score_changed(self, im):
      white_pixels_only = im.copy()
      white_pixels_only[white_pixels_only != 255] = 0
      black_pixels_only = im.copy()
      black_pixels_only[black_pixels_only != 0] = 255
      # If we are still initializing each frame, save the current image
      if self.last_score_images_white[self.last_score_image_index] is None:
        self.last_score_images_white[self.last_score_image_index] = white_pixels_only
        self.last_score_images_black[self.last_score_image_index] = black_pixels_only 
        self.last_score_image_index = (self.last_score_image_index + 1) % self.num_previous_score_images
        return False
      # If we haven't stacked all the images yet, do so
      if self.last_score_images_stacked_white is None:
        self.last_score_images_stacked_white = self.last_score_images_white[0].copy()
        self.last_score_images_stacked_black = self.last_score_images_black[0].copy()
        # For each frame, if a pixel is not white in either the current stack or that frame, its noise, so set that pixel to 0 in the stack
        for im in self.last_score_images_white:
          self.last_score_images_stacked_white[np.logical_or(self.last_score_images_stacked_white == 0, im == 0)] = 0
        # opposite for black
        for im in self.last_score_images_black:
          self.last_score_images_stacked_black[np.logical_or(self.last_score_images_stacked_black == 255, im == 255)] = 255
        return False
      # If the current frame from the game has any pixels that should be white that aren't either the score
      #  changed or there is some noise. Keep track of how many frames in a row have changed, if its more than
      #  some threshold we assume its not noise and the score did change.
      if (white_pixels_only[self.last_score_images_stacked_white == 255] != 255).any() or (black_pixels_only[self.last_score_images_stacked_black == 0] != 0).any():
        self.current_change_in_a_row += 1
        # if its changed enough times in a row, reset everything but the current score, and return True
        if self.current_change_in_a_row == self.require_change_in_a_row:
          save = True
          if save:
            for i, lim in enumerate(self.last_score_images_white):
              Image.fromarray(lim).save(f"./last_score_im_{i}_white_score_{self.current_score}.png", "PNG")
            for i, lim in enumerate(self.last_score_images_black):
              Image.fromarray(lim).save(f"./last_score_im_{i}_black_score_{self.current_score}.png", "PNG")
            Image.fromarray(white_pixels_only).save(f"./current_im_white_score_{self.current_score}.png", "PNG")
            Image.fromarray(black_pixels_only).save(f"./current_im_black_score_{self.current_score}.png", "PNG")
            Image.fromarray(self.last_score_images_stacked_white).save(f"./last_score_im_stacked_white_score_{self.current_score}.png", "PNG")
            Image.fromarray(self.last_score_images_stacked_black).save(f"./last_score_im_stacked_black_score_{self.current_score}.png", "PNG")
          self.last_score_images_white = [None] * self.num_previous_score_images
          self.last_score_images_black = [None] * self.num_previous_score_images
          self.last_score_image_index = 0
          self.last_score_images_stacked_white = None
          self.last_score_images_stacked_black = None
          self.current_change_in_a_row = 0
          self.current_score += 1
          return True
        else:
          return False
      else:
        # if the image completely matches from last frame, we still have the same score, so reset the current count
        self.current_change_in_a_row = 0
      return False


class WindowGetter:
    
    def __init__(self):
        self.screenshot_dims = {"top": 0, "left": 0, "width": 0, "height": 0}
        self.sct = mss.mss()
        self.first_callback = True

    def callback(self, hwnd, extra):
        rect = win32gui.GetWindowRect(hwnd)
        if ("Super Crate Box" in win32gui.GetWindowText(hwnd)):
            super_crate_box_handle = hwnd 
            self.screenshot_dims["top"] = rect[1]
            self.screenshot_dims["left"] = rect[0]
            self.screenshot_dims["width"] = rect[2] - rect[0]
            self.screenshot_dims["height"] = rect[3] - rect[1]
            if self.first_callback:
              win32gui.SetForegroundWindow(super_crate_box_handle)
              self.first_callback = False
    
    def screenshot(self):
        # Grab the data, save as image
        sct_img = self.sct.grab(self.screenshot_dims)
        array = np.asarray(sct_img)
        mss.tools.to_png(sct_img.rgb, sct_img.size, output="images/" + str(time.time()) + ".png")
        return sct_img, array


class ActionSpace:
    
    def __init__(self):
        # self.action_list = [self.left, self.right, self.up, self.left_and_up, 
        #                     self.right_and_up, self.shoot, self.space]
        self.action_list = [self.start_left, self.start_right, self.start_left_and_up, self.start_right_and_up]
        self.left_is_down = False
        self.right_is_down = False
        self.up_is_down = False
        self.threshold = .5
        
    def left(self):
        win32api.keybd_event(0x25, 0, 0, 0)
        time.sleep(.25)
        win32api.keybd_event(0x25, 0, win32con.KEYEVENTF_KEYUP, 0)
        
    def right(self):
        win32api.keybd_event(0x27, 0, 0, 0)
        time.sleep(.25)
        win32api.keybd_event( 0x27, 0, win32con.KEYEVENTF_KEYUP, 0)
            
    def up(self):
        win32api.keybd_event(0x26, 0, 0, 0)
        time.sleep(.25)
        win32api.keybd_event(0x26, 0, win32con.KEYEVENTF_KEYUP, 0)
        
    def left_and_up(self):
        win32api.keybd_event(0x25, 0, 0, 0)
        win32api.keybd_event(0x26, 0, 0, 0)
        time.sleep(.25)
        win32api.keybd_event(0x25, 0, win32con.KEYEVENTF_KEYUP, 0)
        win32api.keybd_event(0x26, 0, win32con.KEYEVENTF_KEYUP, 0)
        
    def right_and_up(self):
        win32api.keybd_event(0x26, 0, 0, 0)
        win32api.keybd_event(0x27, 0, 0, 0)
        time.sleep(.25)
        win32api.keybd_event(0x26, 0, win32con.KEYEVENTF_KEYUP, 0)
        win32api.keybd_event(0x27, 0, win32con.KEYEVENTF_KEYUP, 0)
        
    def shoot(self):
        win32api.keybd_event(0x58, 0, win32con.KEYEVENTF_EXTENDEDKEY, 0)
        win32api.keybd_event(0x58, 0, win32con.KEYEVENTF_KEYUP, 0)
        
    def space(self):
        win32api.keybd_event(0x20, 0, win32con.KEYEVENTF_EXTENDEDKEY, 0)
        win32api.keybd_event(0x20, 0, win32con.KEYEVENTF_KEYUP, 0)    
            
    def get_random_action(self):
        function = np.random.choice(self.action_list)
        function()

    def start_left(self):
      self.let_go_of_right()
      self.let_go_of_up()
      if not self.left_is_down:
        self.left_is_down = True
        win32api.keybd_event(0x25, 0, 0, 0)

    def start_right(self):
      self.let_go_of_left()
      self.let_go_of_up()
      if not self.right_is_down:
        self.right_is_down = True
        win32api.keybd_event(0x27, 0, 0, 0)
    
    def start_left_and_up(self):
      self.let_go_of_right()
      if not self.left_is_down:
        self.left_is_down = True
        win32api.keybd_event(0x25, 0, 0, 0)
      if not self.up_is_down:
        self.up_is_down = True
        win32api.keybd_event(0x26, 0, 0, 0)

    def start_right_and_up(self):
      self.let_go_of_left()
      if not self.right_is_down:
        self.right_is_down = True
        win32api.keybd_event(0x27, 0, 0, 0)
      if not self.up_is_down:
        self.up_is_down = True
        win32api.keybd_event(0x26, 0, 0, 0)

    def let_go_of_left(self):
      if self.left_is_down:
        self.left_is_down = False
        win32api.keybd_event(0x25, 0, win32con.KEYEVENTF_KEYUP, 0)

    def let_go_of_right(self):
      if self.right_is_down:
        self.right_is_down = False
        win32api.keybd_event(0x27, 0, win32con.KEYEVENTF_KEYUP, 0)

    def let_go_of_up(self):
      if self.up_is_down:
        self.up_is_down = False
        win32api.keybd_event(0x26, 0, win32con.KEYEVENTF_KEYUP, 0)
          



    
    
if __name__ == "__main__":
    main()

            

