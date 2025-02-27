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

def merge_presentations(plan, output_file):
    # Create a new presentation
    result_presentation = ppt_instance.Presentations.Add() # create presentation
    insert_index = 1

    for entry in plan['plan']:
        source_presentation = ppt_instance.Presentations.open(os.path.abspath(entry['file']),read_only,has_title,window)
        for slide_index in entry['slides']:
            try:
                source_presentation.Slides(slide_index).Copy()
                result_presentation.Slides.Paste(Index=insert_index)
                insert_index = insert_index + 1
            except Exception as e:
                print(f"An error occurred while copying the slide {slide_index} from {entry['file']} {e}")
                pass
        source_presentation.Close()

    for entry in plan['replace']:
        Replace_String(result_presentation, entry['Slide'], entry['SeekFor'], entry['ReplaceBy'])


    result_presentation.SaveAs(os.path.abspath(output_file))
    result_presentation.Close()

    return

def main(plan_file, output_file):

    plan = load_plan(plan_file)
    merge_presentations(plan, output_file)
    print(f"Generated merged presentation: {output_file}")

    return

ppt_instance = win32com.client.Dispatch('PowerPoint.Application')
read_only = True
has_title = False
window    = False

if __name__ == "__main__":
    parser = argparse.ArgumentParser(description="Name of the plan file (without extension)")
    parser.add_argument('plan', type=str, nargs='?', help='provide the name of a yaml plan file (without extension)')
    args = parser.parse_args()
    if args.plan:
        plan_file = f'{args.plan}.yml'
        output_file = f'{args.plan}.pptx'
        main(plan_file, output_file)
    else:
        print("The plan name is missing")
        print("usage: GenPPT.py [-h] [plan]")

#kills ppt_instance
ppt_instance.Quit()
del ppt_instance
