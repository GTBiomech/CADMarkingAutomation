"""
@author: Dr Gregory Tierney (Email: g.tierney@ulster.ac.uk)
"""

import os
import csv
import win32com.client
from OCC.Core.STEPControl import STEPControl_Reader
from OCC.Core.BRepGProp import brepgprop_VolumeProperties
from OCC.Core.GProp import GProp_GProps
from OCC.Core.BRepGProp import brepgprop_SurfaceProperties
from OCC.Core.IFSelect import IFSelect_RetDone


def export_to_step(par_file_path, output_dir):
    """
    Exports a Solid Edge .par file to a STEP file.
    """
    solid_edge_app = None
    doc = None
    try:
        solid_edge_app = win32com.client.Dispatch('SolidEdge.Application')
        solid_edge_app.Visible = False

        # Open the document
        doc = solid_edge_app.Documents.Open(par_file_path)
        name, _ = os.path.splitext(os.path.basename(par_file_path))
        step_file_path = os.path.join(output_dir, f"{name}.step")

        # Overwrite if the file exists
        if os.path.isfile(step_file_path):
            os.remove(step_file_path)

        # Save as STEP
        doc.SaveAs(step_file_path)
        return step_file_path

    except Exception as e:
        print(f"Error exporting to STEP for {par_file_path}: {e}")
        return None

    finally:
        if doc:
            try:
                doc.Close()
            except Exception as e:
                print(f"Error closing document: {e}")

        if solid_edge_app:
            try:
                solid_edge_app.Quit()
            except Exception as e:
                print(f"Error quitting Solid Edge: {e}")


def clean_up_files(output_dir):
    """
    Deletes any .txt or .log files in the output directory.
    """
    for file in os.listdir(output_dir):
        if file.endswith('.txt') or file.endswith('.log'):
            file_path = os.path.join(output_dir, file)
            try:
                os.remove(file_path)
                print(f"Deleted file: {file_path}")
            except Exception as e:
                print(f"Error deleting file {file_path}: {e}")


def calculate_properties(step_file_path):
    """
    Calculates volume, surface area, and center of gravity from a STEP file.
    """
    try:
        reader = STEPControl_Reader()
        status = reader.ReadFile(step_file_path)

        if status != IFSelect_RetDone:
            print(f"Error reading STEP file: {step_file_path}")
            return None, None, None

        # Transfer and process the STEP file
        reader.TransferRoots()
        shape = reader.OneShape()

        # Calculate volume
        volume_props = GProp_GProps()
        brepgprop_VolumeProperties(shape, volume_props)
        volume = volume_props.Mass()

        # Calculate surface area
        surface_props = GProp_GProps()
        brepgprop_SurfaceProperties(shape, surface_props)
        surface_area = surface_props.Mass()

        # Get the center of gravity
        cg = volume_props.CentreOfMass()
        cg_x, cg_y, cg_z = cg.X(), cg.Y(), cg.Z()

        return volume, surface_area, (cg_x, cg_y, cg_z)

    except Exception as e:
        print(f"Error processing STEP file {step_file_path}: {e}")
        return None, None, None


def extract_expected_values(solution_file_path, output_dir):
    """
    Extracts expected volume, surface area, and CG values from the solution file.
    """
    step_file_path = export_to_step(solution_file_path, output_dir)
    if step_file_path:
        expected_volume, expected_surface_area, expected_cg = calculate_properties(step_file_path)
        if os.path.isfile(step_file_path):
            os.remove(step_file_path)
        clean_up_files(output_dir)
        return expected_volume, expected_surface_area, expected_cg
    return None, None, None


def calculate_mark(value, expected_value):
    """
    Calculates the mark based on percentage difference.
    """
    if value is None or expected_value is None or expected_value == 0:
        return 0
    percentage = abs((value - expected_value) / expected_value) * 100
    if percentage <= 1:
        return 5
    elif percentage <= 20:
        return 4
    elif percentage <= 40:
        return 3
    elif percentage <= 60:
        return 2
    elif percentage <= 80:
        return 1
    return 0


def process_submissions(submissions_folder, output_dir, expected_volume, expected_surface_area, expected_cg):
    """
    Processes all .par files in the submissions folder, calculates properties, and marks them.
    """
    results = []
    submission_count = 0
    total_submissions = len([f for f in os.listdir(submissions_folder) if f.endswith('.par')])

    for file_name in os.listdir(submissions_folder):
        if file_name.endswith('.par'):
            student_id, _ = os.path.splitext(file_name)
            par_file_path = os.path.join(submissions_folder, file_name)

            retry_count = 3
            step_file_path = None
            for _ in range(retry_count):
                step_file_path = export_to_step(par_file_path, output_dir)
                if step_file_path:
                    break

            if not step_file_path:
                print(f"Failed to export {file_name} to STEP after {retry_count} attempts.")
                continue

            # Calculate properties and mark the submission
            volume, surface_area, cg = calculate_properties(step_file_path)

            if volume is None or surface_area is None or cg is None:
                print(f"Error calculating properties for {file_name}. Skipping.")
                continue

            # Calculate marks
            volume_mark = calculate_mark(volume, expected_volume)
            surface_area_mark = calculate_mark(surface_area, expected_surface_area)

            cg_x_mark = calculate_mark(cg[0], expected_cg[0]) / 3  # Maximum of 1.6666
            cg_y_mark = calculate_mark(cg[1], expected_cg[1]) / 3  # Maximum of 1.6666
            cg_z_mark = calculate_mark(cg[2], expected_cg[2]) / 3  # Maximum of 1.6666
            cg_mark = cg_x_mark + cg_y_mark + cg_z_mark  # Sum up to a max of 5

            results.append({
                'StudentID': student_id,
                'Volume': volume,
                'SurfaceArea': surface_area,
                'CenterOfGravity': cg,
                'VolumeMark': volume_mark,
                'SurfaceAreaMark': surface_area_mark,
                'CGMark': cg_mark
            })

            if os.path.isfile(step_file_path):
                os.remove(step_file_path)

            # Real-time progress update
            submission_count += 1
            print(f"Processed {submission_count}/{total_submissions} submissions.")

    clean_up_files(output_dir)
    return results


def save_results_to_csv(results, output_file):
    """
    Saves the results to a CSV file.
    
    """
    with open(output_file, mode='w', newline='') as csv_file:
        fieldnames = ['StudentID', 'Volume', 'SurfaceArea', 'CenterOfGravity', 'VolumeMark', 'SurfaceAreaMark', 'CGMark']
        writer = csv.DictWriter(csv_file, fieldnames=fieldnames)

        writer.writeheader()
        for result in results:
            writer.writerow({
                'StudentID': result['StudentID'],
                'Volume': f"{result['Volume']:.3f}" if isinstance(result['Volume'], (float, int)) else result['Volume'],
                'SurfaceArea': f"{result['SurfaceArea']:.3f}" if isinstance(result['SurfaceArea'], (float, int)) else result['SurfaceArea'],
                'CenterOfGravity': (
                    f"({result['CenterOfGravity'][0]:.3f}, "
                    f"{result['CenterOfGravity'][1]:.3f}, "
                    f"{result['CenterOfGravity'][2]:.3f})"
                ) if isinstance(result['CenterOfGravity'], tuple) else result['CenterOfGravity'],
                'VolumeMark': f"{result['VolumeMark']:.2f}",
                'SurfaceAreaMark': f"{result['SurfaceAreaMark']:.2f}",
                'CGMark': f"{result['CGMark']:.2f}",
            })

"""
# Example usage (Replace C:\Users\... with own path)
solution_file_path = r'C:\Users\...'
submissions_folder = r'C:\Users\...'
output_dir = r'C:\Users\...'
output_csv = os.path.join(output_dir, 'submission_results.csv')

expected_volume, expected_surface_area, expected_cg = extract_expected_values(solution_file_path, output_dir)
submission_results = process_submissions(submissions_folder, output_dir, expected_volume, expected_surface_area, expected_cg)
save_results_to_csv(submission_results, output_csv)

print(f"Results saved to {output_csv}")
"""
