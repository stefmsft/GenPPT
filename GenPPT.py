import yaml
import win32com.client
import os
import argparse

def load_plan(plan_file):
    with open(plan_file, 'r') as file:
        return yaml.safe_load(file)

def Replace_String(Pres, SlideNb, SeekFor, ReplaceBy):

    fs = Pres.Slides(SlideNb)
    target_string = SeekFor
    replacement_string = ReplaceBy
    # Iterate through the shapes on the first slide
    for shape in fs.Shapes:
        if shape.HasTextFrame:
            if shape.TextFrame.HasText:
                text = shape.TextFrame.TextRange.Text
                if target_string in text:
                    # Replace the target string with the replacement string
                    shape.TextFrame.TextRange.Text = text.replace(target_string, replacement_string)

    return

def GetSectionAndSetsFromProfile(plan,targetprofile):

    listofsections = []
    listofsets = []
    for profile in plan['Profiles']:
        try:
            sectioname = list(profile.keys())[0]
            if sectioname.lower() == targetprofile.lower():
                if "Sections" in profile.keys():
                    listofsections = profile['Sections']
                if "LabelSet" in profile.keys():
                    listofsets = profile['LabelSet']
                break
        except Exception as e:
            print(f"An error occurred while checking profile {profile} {e}")
            return
    return listofsections,listofsets

def GetTouchUpListFromPlan(plan):

    listofTouchUp = []

    for mod in plan['TouchUp']:
        try:

            if "Slide" in mod.keys():
                mslide = mod['Slide']
            if "SeekFor" in mod.keys():
                mseek = mod['SeekFor']
            if "ReplaceBy" in mod.keys():
                mreplace = mod['ReplaceBy']
            listofTouchUp.append({"Slide":mslide,"SeekFor":mseek,"ReplaceBy":mreplace})

        except Exception as e:
            print(f"An error occurred while checking profile {profile} {e}")
            return
    return listofTouchUp

def GetSectionFromName(plan,sectionname):

    section = None

    index = next((i for i, d in enumerate(plan['Sections']) if list(d.keys())[0] == sectionname), -1)
    if index != -1:
        section = plan['Sections'][index]

    return section

def AddToList(List,sdic):

    if "Reffile" in sdic.keys():
        RefFile = sdic['Reffile']
        if "slides" in sdic.keys():
            SlideSet = sdic['slides']
            List.append({"Reffile":RefFile,"slides":SlideSet})

    return List

def GetTargetSlidesFromPlanSectionsAndSets(plan,Sections,Setslist):

    listofslides = []
    SectionsAvailable = [list(d.keys())[0] for d in plan['Sections']]

    for section in Sections:
        try:
            if section in SectionsAvailable:
                sectiontoget = GetSectionFromName(plan,section)
                if sectiontoget != None:
                    if len(sectiontoget.keys()) == 1:
                        for sets in sectiontoget[section]:
                            if (sets['Set'] in Setslist) or (sets['Set'] == None):
                                AddToList(listofslides,sets)
                                break
                    else:
                        AddToList(listofslides,sectiontoget)
                else:
                    print(f"An error occurred while checking profile {section} {e}")
        except Exception as e:
            print(f"An error occurred while checking profile {section} {e}")
            return

    return listofslides

def ApplyTouchUp(Presentation,TouchUpList):

    for mod in TouchUpList:
        Replace_String(Presentation, mod['Slide'], mod['SeekFor'], mod['ReplaceBy'])

    return Presentation

def GatherSlides(TargetSlides):
    # Create a new presentation
    result_presentation = ppt_instance.Presentations.Add() # create presentation
    insert_index = 1

    for entry in TargetSlides:
        source_presentation = ppt_instance.Presentations.open(os.path.abspath(entry['Reffile']),read_only,has_title,window)
        for slide_index in entry['slides']:
            try:
                source_presentation.Slides(slide_index).Copy()
                result_presentation.Slides.Paste(Index=insert_index)
                insert_index = insert_index + 1
            except Exception as e:
                print(f"An error occurred while copying the slide {slide_index} from {entry['file']} {e}")
                pass
        source_presentation.Close()

    return result_presentation

def main(plan_file, output_file, profil):

    plan = load_plan(plan_file)
    Sections,Sets = GetSectionAndSetsFromProfile(plan,profil)
    if len(Sections) > 0:
        TargetSlides = GetTargetSlidesFromPlanSectionsAndSets(plan,Sections,Sets)
        Presentation = GatherSlides(TargetSlides)
        TouchUpList = GetTouchUpListFromPlan(plan)
        if len(TouchUpList) > 0:
            Presentation = ApplyTouchUp(Presentation,TouchUpList)

    NbSlides = Presentation.Slides.Count
    Presentation.SaveAs(os.path.abspath(output_file))
    Presentation.Close()

#    merge_presentations(plan, output_file)
    print(f"Generated merged presentation: {output_file}")
    print(f'With {NbSlides} Slides for a presentation time of arround {(NbSlides-2)*3} minutes')

    return

ppt_instance = win32com.client.Dispatch('PowerPoint.Application')
read_only = True
has_title = False
window    = False

plan_file = ""
output_file = ""
ProfileRequested = "Default"

if __name__ == "__main__":
    parser = argparse.ArgumentParser(description="Name of the plan file (without extension)")
    parser.add_argument('plan', type=str, nargs='?', help='provide the name of a yaml plan file (without extension)')
    parser.add_argument('-p', '--profile', type=str, help='The profile name to use.')
    args = parser.parse_args()
    if args.plan:
        plan_file = f'{args.plan}.yml'
        output_file = f'{args.plan}.pptx'
    else:
        print("The plan name is missing")
        print("usage: GenPPT.py [-h] planName [-profile profilname]")
        exit()
    if args.profile:
        ProfileRequested = args.profile

main(plan_file, output_file, ProfileRequested)

#kills ppt_instance
ppt_instance.Quit()
del ppt_instance
