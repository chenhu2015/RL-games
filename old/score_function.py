import time
import numpy as np
import win32api
import win32com.client
import win32con
import win32gui
from PIL import ImageOps, Image
import d3dshot
import matplotlib.pyplot as plt
d = d3dshot.create()
super_crate_box_handle = None
screenshot_bounds = None
trimming = (8, 31, -8, -8)

def callback(hwnd, extra):
  rect = win32gui.GetWindowRect(hwnd)
  x = rect[0]
  y = rect[1]
  w = rect[2] - x
  h = rect[3] - y
  if ("Super Crate Box" in win32gui.GetWindowText(hwnd)):
    super_crate_box_handle = hwnd
    print("Window %s:" % win32gui.GetWindowText(hwnd))
    print("\tLocation: (%d, %d)" %(x, y))
    print("\t    Size: (%d, %d)" % (w, h))
    global screenshot_bounds
    screenshot_bounds = (x + trimming[0],
                         y + trimming[1],
                         x + w + trimming[2], 
                         y + h + trimming[3])
    win32gui.SetForegroundWindow(super_crate_box_handle)



# X then Y
coords_of_pixel_in_G_of_GameOver = (77, 72)
coords_of_pixel_under_G_of_GameOver = (77, 77)
coords_of_pixel_in_R = (166, 75)
def is_image_of_game_over(im):
  a = coords_of_pixel_in_G_of_GameOver
  b = coords_of_pixel_under_G_of_GameOver
  c = coords_of_pixel_in_R
  return (im[a[1]][a[0]] == 255 and im[b[1]][b[0]] == 0 and im[c[1]][c[0]] == 255)

# SCORING STUFF ###############################
score_bounding_box = (100, 7, 139, 22) 
num_previous_score_images = 5
last_score_images_white = [None] * num_previous_score_images
last_score_images_black = [None] * num_previous_score_images
last_score_image_index = 0
last_score_images_stacked_white = None
last_score_images_stacked_black = None
require_change_in_a_row = 5
current_change_in_a_row = 0
current_score = 0

# Preprocess the image, it should be passed cropped. All it does right now is turn it black and white
# copy it, and return it.
def score_preprocess(im):
  processed = im.copy()
  processed[processed != 255] = 0
  return processed

# Used to reset all the global variable to their initial state
def reset_scoring_function():
  global last_score_images_white, last_score_image_index, last_score_images_stacked_white
  global last_score_images_black, last_score_images_stacked_black
  global require_change_in_a_row, current_change_in_a_row, current_score
  last_score_images_white = [None] * num_previous_score_images
  last_score_images_black = [None] * num_previous_score_images
  last_score_image_index = 0
  last_score_images_stacked_white = None
  last_score_images_stacked_black = None
  current_change_in_a_row = 0
  current_score = 0

# Used to check if the score has changed. Need a few frames in a row to have been different before it is registered
def score_changed(im):
  global last_score_images_white, last_score_image_index, last_score_images_stacked_white
  global last_score_images_black, last_score_images_stacked_black
  global require_change_in_a_row, current_change_in_a_row, current_score
  white_pixels_only = im.copy()
  white_pixels_only[white_pixels_only != 255] = 0
  black_pixels_only = im.copy()
  black_pixels_only[black_pixels_only != 0] = 255
  # If we are still initializing each frame, save the current image
  if last_score_images_white[last_score_image_index] is None:
    last_score_images_white[last_score_image_index] = white_pixels_only
    last_score_images_black[last_score_image_index] = black_pixels_only
    last_score_image_index = (last_score_image_index + 1) % num_previous_score_images
    return False
  # If we haven't stacked all the images yet, do so
  if last_score_images_stacked_white is None:
    last_score_images_stacked_white = last_score_images_white[0].copy()
    last_score_images_stacked_black = last_score_images_black[0].copy()
    # For each frame, if a pixel is not white in either the current stack or that frame, its noise, so set that pixel to 0 in the stack
    for im in last_score_images_white:
      last_score_images_stacked_white[np.logical_or(last_score_images_stacked_white == 0, im == 0)] = 0
    # opposite for black
    for im in last_score_images_black:
      last_score_images_stacked_black[np.logical_or(last_score_images_stacked_black == 255, im == 255)] = 255
    return False
  # If the current frame from the game has any pixels that should be white that aren't either the score
  #  changed or there is some noise. Keep track of how many frames in a row have changed, if its more than
  #  some threshold we assume its not noise and the score did change.
  if (white_pixels_only[last_score_images_stacked_white == 255] != 255).any() or (black_pixels_only[last_score_images_stacked_black == 0] != 0).any():
    current_change_in_a_row += 1
    # if its changed enough times in a row, reset everything but the current score, and return True
    if current_change_in_a_row == require_change_in_a_row:
      save = False
      if save:
        for i, lim in enumerate(last_score_images_white):
          Image.fromarray(lim).save(f"./last_score_im_{i}_white_score_{current_score}.png", "PNG")
        for i, lim in enumerate(last_score_images_black):
          Image.fromarray(lim).save(f"./last_score_im_{i}_black_score_{current_score}.png", "PNG")
        Image.fromarray(white_pixels_only).save(f"./current_im_white_score_{current_score}.png", "PNG")
        Image.fromarray(black_pixels_only).save(f"./current_im_black_score_{current_score}.png", "PNG")
        Image.fromarray(last_score_images_stacked_white).save(f"./last_score_im_stacked_white_score_{current_score}.png", "PNG")
        Image.fromarray(last_score_images_stacked_black).save(f"./last_score_im_stacked_black_score_{current_score}.png", "PNG")
      last_score_images_white = [None] * num_previous_score_images
      last_score_images_black = [None] * num_previous_score_images
      last_score_image_index = 0
      last_score_images_stacked_white = None
      last_score_images_stacked_black = None
      current_change_in_a_row = 0
      current_score += 1
      return True
    else:
      return False
  else:
    # if the image completely matches from last frame, we still have the same score, so reset the current count
    current_change_in_a_row = 0
  return False


# function to detect when the screen shakes, so we can selectively not deal with it.
# list of pixel values for when the screen isn't shaking. If any of these are not
# valid, the screen is shaking
screenshake_in_a_row = 3 # when screenshake is detected, assume it lasts for 3 frames
expected_pixel_map = {(0,0):105,(0,50):105,(50,0):105, (1,1):91}
def is_image_with_screen_shake(im):
  global screenshake_in_a_row
  if (screenshake_in_a_row > 0):
    screenshake_in_a_row -= 1
    return True
  for key, val in expected_pixel_map.items():
    if im[key[0], key[1]] != val:
      screenshake_in_a_row = 3
      return True
  return False

if __name__ == "__main__":
  win32gui.EnumWindows(callback, None)
  shell = win32com.client.Dispatch("WScript.Shell")
  while True:
    im = d.screenshot()
    im = ImageOps.grayscale(im)
    im = im.crop(screenshot_bounds)
    if not is_image_with_screen_shake(np.array(im)):
      score_im = im.crop(score_bounding_box)
      if score_changed(np.array(score_im)):
        print("Score: ", current_score)
    if is_image_of_game_over(np.array(im)):
      shell.SendKeys("Z")
      reset_scoring_function()
      print("Score: ", current_score)
    #im.show()
    #break
# Key codes here: https://docs.microsoft.com/en-us/dotnet/api/system.windows.forms.sendkeys?view=netcore-3.1




# super crate box is 240 x 160 on smallest setting

