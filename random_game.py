import random
import sys
import pygame
from pygame.locals import *
# global Variables for the game

FPS = 32
SCRWIDTH = 289
SCRHIGHT = 511
SCREEN = pygame.dislay.set_mode((SCRWIDTH,SCRHIGHT))
GROUNDY = SCRHIGHT * 0.8
GAME_SPRITES = {}
GAME_SOUNDS = {}
#PLAYER= ''
#BACKGROUND
#PIPE

def welcomeScreen():
    #shows welcome images on screen
    playerx =int(SCREENWIDTH / 5)
    playery =