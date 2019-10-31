import math,cairo
import datetime

from dateutil.relativedelta import *

class GanttChart:

    def __init__(self, date_project_start):
        fn_cairo = "gantt.png"
        self.date_project_start=date_project_start

        self.width, self.height = 768,768
        #        surface = cairo.PSSurface(fn_cairo, self.width, self.height)

        self.surface = cairo.ImageSurface(cairo.FORMAT_ARGB32, self.width, self.height)

        self.ctx = cairo.Context (self.surface)
        self.drawGrid()

    def draw(self, fn_out_png):  #draw it to the file and finish
        self.ctx.show_page()

        self.surface.write_to_png(fn_out_png)
        print("done draw")

    def t2x(self, t):   #time to x coord
        return int( 0.0 + self.width*t/12.    )

    def r2y(self, r):  #row to y coord
        return int( 0.0 + self.height*r/38.    )  #CHANGE YSCALE HERE*****


    def drawTask(self, i_row_in, t_start, t_end, name):
        i_row=i_row_in+1  #HACK move everythign down 1 line to make space at top
        ctx=self.ctx
        xstart=self.t2x(t_start)
        xend=self.t2x(t_end)
        ystart=self.r2y(i_row)
        yend=self.r2y(i_row+1)
        w=xend-xstart
        h=yend-ystart

        ctx.set_source_rgb(.5,.5,1)
        ctx.rectangle(xstart,ystart,w,h)
        ctx.fill()
        ctx.select_font_face("Sans", cairo.FONT_SLANT_NORMAL, cairo.FONT_WEIGHT_BOLD)
        ctx.set_source_rgb(0,0,.3)
        ctx.set_font_size(12)
        ctx.move_to(xstart, yend)
        ctx.show_text(name)

    def drawGrid(self):
        ctx=self.ctx
#        date_project_start = datetime.datetime.strptime('2019-04-01_00:09:00.00', "%Y-%m-%d_%H:%M:%S.%f" )
        for q in range(1,10):
            x = self.t2x(q)
            ystart=0
            yend=self.height

            #quarter name text
            qname="Q"+str(q)
            date_quarter_start = self.date_project_start+ relativedelta(months=+((q-1)*3))

            ctx.set_source_rgb(0,0,0)
            ctx.move_to(x, ystart)
            ctx.line_to(x, yend)
            ctx.stroke()
            ctx.select_font_face("Sans", cairo.FONT_SLANT_NORMAL, cairo.FONT_WEIGHT_BOLD)
            ctx.set_source_rgb(0,0,0)
            ctx.set_font_size(12)
            ctx.move_to(x, ystart+10)   
            ctx.show_text(qname)
            ctx.move_to(x, ystart+20)   
            qdate=date_quarter_start.strftime("%Y-%m")   #TODO convert to actual dates
            ctx.show_text(qdate)


if __name__=="__main__":
    g = GanttChart()
    g.drawTask( i_row=1, t_start=1, t_end=3, name="D1.1")
    g.drawTask( i_row=2, t_start=2, t_end=5, name="D1.2")
    g.draw("test.png")

