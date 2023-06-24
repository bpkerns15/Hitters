import pandas as pd
import numpy as np
from pptx import Presentation
from pptx.util import Pt
from pptx.util import Inches
from pptx.oxml.xmlchemy import OxmlElement
from pptx.util import Cm
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import PP_ALIGN
from pptx.dml.color import RGBColor
from pptx.enum.text import MSO_ANCHOR
import six
import copy
import seaborn as sns
import matplotlib.pyplot as plt
from PIL import Image
import matplotlib.patches as patches
import os
import math
from scipy import interpolate
import matplotlib.lines as lines
from tkinter import filedialog
from scipy.spatial import QhullError

opponent = 'KAM_NOR'

def get_csv():
    global opponent
    #csv_file = 'CSVs//OregonHitters.csv'
    csv_file = 'Kamploops.csv'
    csv_df = pd.read_csv(csv_file)
    csv_df = csv_df.drop(csv_df[csv_df.BatterTeam != opponent].index)
    return csv_df

def to_player_df(csv_df, name):
    player_df = csv_df[(csv_df['Batter'] == name)]
    return player_df

def player_df_to_side_df(player_df, side):
    side_df = player_df[(player_df['PitcherThrows'] == side)]
    return side_df

def data_frame_for_damage_chart(pitch_df):
    damage_df = pitch_df

    #Remove unnesassary columns
    remove_list = damage_df.columns.values.tolist()
    remove_list.remove("Batter")
    remove_list.remove("PlateLocHeight")
    remove_list.remove("PlateLocSide")
    remove_list.remove("BatterSide")
    remove_list.remove("ExitSpeed")
    damage_df = damage_df.drop(remove_list, axis=1)

    #Removes Rows where EV is nan
    drop_index = []
    for index, row in damage_df.iterrows():
        if math.isnan(row['ExitSpeed']) == True:
            drop_index.append(index)
    damage_df = damage_df.drop(drop_index)

    #Removes Rows where EV < 60
    damage_df = damage_df.drop(damage_df[damage_df.ExitSpeed < 55.0].index)

    #Switches from Pitcher view to catcher view
    #damage_df['PlateLocSide'] = (damage_df['PlateLocSide'] * -1)

    return damage_df

def damage_chart(damage_df, name, side):
    x = []
    y = []
    array = []
    print(damage_df)
    for i in range(10):
        row = []
        for j in range(8):
            temp_df = damage_df
            top_limit = 4.267222-0.35819444*i
            bottom_limit = 4.267222-0.35819444*(i+1)
            left_limit = -1.4166667 + 0.35416667*j
            right_limit = -1.416667 + 0.35416667*(j+1)

            #drops data from df not in cell needed
            too_left = temp_df[(temp_df['PlateLocSide'] < left_limit)].index
            temp_df = temp_df.drop(too_left)
            too_right = temp_df[(temp_df['PlateLocSide'] > right_limit)].index
            temp_df = temp_df.drop(too_right)
            too_high = temp_df[(temp_df['PlateLocHeight'] > top_limit)].index
            temp_df = temp_df.drop(too_high)
            too_low = temp_df[(temp_df['PlateLocHeight'] < bottom_limit)].index
            temp_df = temp_df.drop(too_low)

            avg_ev_for_zone = round(temp_df["ExitSpeed"].mean(),1)
            
                

            row.append(avg_ev_for_zone)

        array.append(row)


    df = pd.DataFrame(array)
    print(df)

    try:
        df.to_numpy()
        x = np.arange(0, df.shape[1])
        y = np.arange(0, df.shape[0])
        #mask invalid values
        df = np.ma.masked_invalid(df)
        xx, yy = np.meshgrid(x, y)
        #get only the valid values
        x1 = xx[~df.mask]
        y1 = yy[~df.mask]
        newarr = df[~df.mask]

        GD1 = interpolate.griddata((x1, y1), newarr.ravel(),
                                (xx, yy),
                                    method='linear', fill_value=55.0)
        df = pd.DataFrame(GD1)
    except (ValueError,QhullError) as e:
        df = pd.DataFrame(array)
        df = df.fillna(55.0)


    
    print(df)

    fig, ax = plt.subplots()
    rect = patches.Rectangle((1.5, 1.5), 4, 6, linewidth=1, edgecolor='black', facecolor='none')
    ax.add_patch(rect)
    horizontal_line1 = lines.Line2D([1.5, 5.5], [5.5, 5.5], linewidth=1, color='black')
    horizontal_line2 = lines.Line2D([1.5, 5.5], [3.5, 3.5], linewidth=1, color='black')
    vertical_line1 = lines.Line2D([2.833333, 2.833333], [1.5, 7.5], linewidth=1, color='black')
    vertical_line2 = lines.Line2D([4.166667, 4.166667], [1.5, 7.5], linewidth=1, color='black')
    ax.add_line(horizontal_line1)
    ax.add_line(horizontal_line2)
    ax.add_line(vertical_line1)
    ax.add_line(vertical_line2)

    fig = plt.imshow(df, cmap = 'jet',vmin=55,vmax=95,  interpolation='bicubic')
    #cbar = plt.colorbar(fig)
    fig.axes.get_xaxis().set_visible(False)
    fig.axes.get_yaxis().set_visible(False)
    #plt.title("Exit Velo Heat Map")
    newpath = os.path.join("Plots", name)
    if not os.path.exists(newpath):
        os.makedirs(newpath)
    #saves plot in folder
    plt.axis('off')
    plt.savefig(os.path.join("Plots", name, side + 'HP_heatmap.png'), bbox_inches='tight', pad_inches = 0)

    # load the image
    img = Image.open(os.path.join("Plots", name, side + 'HP_heatmap.png'))

    # get the color you want to make transparent
    color_to_replace = (0, 0, 127)  # in this example, we want to replace all shades of blue

    # make the color transparent
    img = img.convert("RGBA")
    data = img.getdata()

    new_data = []
    for item in data:
        if item[0] == color_to_replace[0] and item[1] == color_to_replace[1] and item[2] == color_to_replace[2]:
            new_data.append((255, 255, 255, 0))  # replace the color with a fully transparent color
        else:
            new_data.append(item)

    img.putdata(new_data)

    # save the modified image
    img.save(os.path.join("Plots", name, side + 'HP_heatmap.png'))

    return

def make_presentation():
    prs = Presentation("Template.pptx")
    prs.slide_width = Inches(10.5)
    prs.slide_height = Inches(8.115)
    #slide = prs.slides.add_slide(prs.slide_layouts[5])

    return prs

def ev_calculations(player_df, lhp_df, rhp_df):

    lhp_df = lhp_df[lhp_df['ExitSpeed'] >= 55.0]
    rhp_df = rhp_df[rhp_df['ExitSpeed'] >= 55.0]
    player_df = player_df[player_df['ExitSpeed'] >= 55.0]
    lhp_df.dropna(subset=['ExitSpeed'], inplace=True)
    rhp_df.dropna(subset=['ExitSpeed'], inplace=True)
    player_df.dropna(subset=['ExitSpeed'], inplace=True)
    

    max_exit_speed = str(round(player_df['ExitSpeed'].max(),1))
    average_exit_speed = str(round(player_df['ExitSpeed'].mean(),1))
    average_exit_speed_rhp = str(round(rhp_df['ExitSpeed'].mean(),1))
    average_exit_speed_lhp = str(round(lhp_df['ExitSpeed'].mean(),1))

    over_90_lhp = lhp_df[lhp_df['ExitSpeed'] >= 90.0]
    over_90_rhp = rhp_df[rhp_df['ExitSpeed'] >= 90.0]

    hard_hit_rhp = str(round((over_90_rhp.shape[0]/rhp_df.shape[0])*100,1))
    hard_hit_lhp = str(round((over_90_lhp.shape[0]/lhp_df.shape[0])*100,1))

    return_list = [max_exit_speed,average_exit_speed_lhp,average_exit_speed,average_exit_speed_rhp,hard_hit_lhp,hard_hit_rhp]

    return return_list



def presentation(name, prs, ev_stats, i):

    first_slide_id = prs.slides[0].slide_id
    slide = prs.slides.get(first_slide_id)
    name1 = name.split()
    full_name = name1[-1] + ' ' + name1[0]
    full_name = full_name[:-1]
    title = slide.shapes.title
    title.text = full_name
    title.text_frame.paragraphs[0].font.size = Pt(27)
    title.text_frame.paragraphs[0].font.name = 'Beaver Bold'
    title.text_frame.paragraphs[0].alignment = PP_ALIGN.CENTER
    title.vertical_anchor = MSO_ANCHOR.MIDDLE

    rhp = os.path.join("Plots", name, 'Right' + 'HP_heatmap.png')
    lhp = os.path.join("Plots", name, 'Left' + 'HP_heatmap.png')
    shape = slide.shapes.add_picture(lhp,Inches(1.54),Inches(2.42), width=Inches(1.73), height=Inches(2.17))
    shape = slide.shapes.add_picture(rhp,Inches(5.09),Inches(2.42), width=Inches(1.73), height=Inches(2.17))

    #EVs RHP Hard Hit%
    x, y, cx, cy = Inches(6.41), Inches(6.73), Inches(0.99), Inches(1.32)
    shape = slide.shapes.add_table((2), 1, x, y, cx, cy)
    table = shape.table
    tbl =  shape._element.graphic.graphicData.tbl
    style_id = '{2D5ABB26-0587-4C30-8999-92F81FD0307C}'
    tbl[0][-1].text = style_id

    cell = table.cell(0, 0)
    cell.text = 'Hard Hit % RHP'
    cell.text_frame.paragraphs[0].font.size = Pt(14)
    cell.text_frame.paragraphs[0].font.name = 'Adobe Heiti Std R'
    cell.text_frame.paragraphs[0].alignment = PP_ALIGN.CENTER
    cell.vertical_anchor = MSO_ANCHOR.MIDDLE

    cell = table.cell(1, 0)
    cell.text = ev_stats[5]
    cell.text_frame.paragraphs[0].font.size = Pt(14)
    cell.text_frame.paragraphs[0].font.name = 'Adobe Heiti Std R'
    cell.text_frame.paragraphs[0].alignment = PP_ALIGN.CENTER
    cell.vertical_anchor = MSO_ANCHOR.MIDDLE

    #LHP Hard Hit%
    x, y, cx, cy = Inches(2.79), Inches(6.73), Inches(0.99), Inches(1.32)
    shape = slide.shapes.add_table((2), 1, x, y, cx, cy)
    table = shape.table
    tbl =  shape._element.graphic.graphicData.tbl
    style_id = '{2D5ABB26-0587-4C30-8999-92F81FD0307C}'
    tbl[0][-1].text = style_id

    cell = table.cell(0, 0)
    cell.text = 'Hard Hit % LHP'
    cell.text_frame.paragraphs[0].font.size = Pt(14)
    cell.text_frame.paragraphs[0].font.name = 'Adobe Heiti Std R'
    cell.text_frame.paragraphs[0].alignment = PP_ALIGN.CENTER
    cell.vertical_anchor = MSO_ANCHOR.MIDDLE

    cell = table.cell(1, 0)
    cell.text = ev_stats[4]
    cell.text_frame.paragraphs[0].font.size = Pt(14)
    cell.text_frame.paragraphs[0].font.name = 'Adobe Heiti Std R'
    cell.text_frame.paragraphs[0].alignment = PP_ALIGN.CENTER
    cell.vertical_anchor = MSO_ANCHOR.MIDDLE

    #ev avgs
    x, y, cx, cy = Inches(4.2), Inches(6.73), Inches(0.99), Inches(2.14)
    shape = slide.shapes.add_table((2), 3, x, y, cx, cy)
    table = shape.table
    tbl =  shape._element.graphic.graphicData.tbl
    style_id = '{D7AC3CCA-C797-4891-BE02-D94E43425B78}'
    tbl[0][-1].text = style_id

    text_labels = ['LHP EV','AVG EV','RHP EV']
    for i in range(3):
        cell = table.cell(0, i)
        cell.text = text_labels[i]
        cell.text_frame.paragraphs[0].font.size = Pt(14)
        cell.text_frame.paragraphs[0].font.name = 'Adobe Heiti Std R'
        cell.text_frame.paragraphs[0].alignment = PP_ALIGN.CENTER
        cell.vertical_anchor = MSO_ANCHOR.MIDDLE
        cell.fill.solid()
        cell.fill.fore_color.rgb = RGBColor(152, 1, 46)
        cell.text_frame.paragraphs[0].font.color.rgb = RGBColor(255, 248, 221)

        cell = table.cell(1, i)
        cell.text = ev_stats[1+i]
        cell.text_frame.paragraphs[0].font.size = Pt(14)
        cell.text_frame.paragraphs[0].font.name = 'Adobe Heiti Std R'
        cell.text_frame.paragraphs[0].alignment = PP_ALIGN.CENTER
        cell.vertical_anchor = MSO_ANCHOR.MIDDLE
        cell.fill.solid()
        cell.fill.fore_color.rgb = RGBColor(255, 255, 255)
        cell.text_frame.paragraphs[0].font.color.rgb = RGBColor(0, 0, 0)

    newpath = os.path.join("FinalPres", opponent)
    if not os.path.exists(newpath):
        os.makedirs(newpath)

    prs.save(os.path.join("FinalPres", opponent, opponent + '.pptx'))

    return prs

def main():
    csv = get_csv()
    prs = make_presentation()
    names = csv["Batter"].unique()
    for i in range(1):
        player_df = to_player_df(csv,names[i])
        rhp_df = player_df_to_side_df(player_df,'Right')
        lhp_df = player_df_to_side_df(player_df,'Left')
        ev_stats = ev_calculations(player_df,lhp_df,rhp_df)
        lhp_df_dc = data_frame_for_damage_chart(rhp_df)
        rhp_df_dc = data_frame_for_damage_chart(lhp_df)
        damage_chart(lhp_df_dc,names[i],'Left')
        damage_chart(rhp_df_dc,names[i],'Right')
        prs = presentation(names[i],prs,ev_stats,i)
        
main()




