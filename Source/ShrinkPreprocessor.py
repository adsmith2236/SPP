import re
import numpy as np
from tkinter import filedialog
from tkinter import *
from tkinter import messagebox as mbox

BOHR_TO_ANG = 0.529177

molecule = []
molecule_bonds = set()
angles = set()
linear_bends = set()
torsions = set()
plane_bends = set()


def reset_molecule():
    molecule.clear()
    molecule_bonds.clear()
    angles.clear()
    linear_bends.clear()
    torsions.clear()
    plane_bends.clear()


def identify_input(filename):
    reset_molecule()
    file = open(filename, "r")
    force_constants = ""
    if filename.endswith(".fchk"):
        force_constants = read_gaussian_chck(file)
    else:
        line = file.readline()
        while line != "":
            if "NWChem" in line:
                force_constants = read_nesi(file)
            elif "Gaussian" in line:
                force_constants = read_gaussian(file)
            line = file.readline()
    file.close()
    return force_constants


def read_gaussian(file):
    force_constants = ""
    line = file.readline()
    while line != "":
        if "Standard orientation:" in line.strip():
            molecule.clear()
            for i in range(4):
                file.readline()
            line = file.readline().strip()
            while not line.startswith("-"):
                atomic_parameters = re.split("(?<! ) ", line)
                for i in range(len(atomic_parameters)):
                    atomic_parameters[i] = atomic_parameters[i].strip()
                if len(molecule) == 0:
                    # blank atom to align indexes with input values
                    molecule.append(Atom())
                molecule.append(Atom(int(atomic_parameters[1]), np.array(
                    [float(atomic_parameters[3]), float(atomic_parameters[4]), float(atomic_parameters[5])])))
                line = file.readline().strip()
        elif "Optimized Parameters" in line.strip() and len(molecule) != 0:
            for atom in molecule:
                atom.clear_bonds()
            for i in range(4):
                file.readline()
            line = file.readline().strip()
            while "R" in line:
                bonded_atoms = re.split("(?<! ) ", line)[2].strip()
                atom_1 = int(bonded_atoms.split(",")[0].strip("R("))
                atom_2 = int(bonded_atoms.split(",")[1].strip(")"))
                molecule[atom_1].add_bond(atom_2)
                molecule[atom_2].add_bond(atom_1)
                line = file.readline().strip()
        elif "Frc consts" in line.strip():
            column = 1
            while "Thermochemistry" not in line.strip():
                if "Frc consts" in line.strip():
                    values = re.split("(?<! ) ", line.strip())
                    for value in values[3:]:
                        force_constant = float(value.strip())
                        if force_constant > 0:
                            force_constants += " "
                        force_constants += " {:.8e}".format(force_constant)
                        if column == 5:
                            force_constants += "\n"
                            column = 0
                        column += 1
                line = file.readline()
        line = file.readline()
    return force_constants


def read_gaussian_chck(file):
    line = file.readline()
    force_constants = ""
    while line != "":
        if "Atomic numbers" in line.strip():
            line = file.readline().strip()
            while not line.startswith("N"):
                atomic_numbers = re.split("(?<! ) ", line)
                for atomic_number in atomic_numbers:
                    if len(molecule) == 0:
                        # blank atom to align indexes with input values
                        molecule.append(Atom())
                    molecule.append(Atom(int(atomic_number.strip())))
                line = file.readline().strip()
        elif "Current cartesian coordinates" in line.strip():
            line = file.readline().strip()
            atom_index = 1
            x = None
            y = None
            while not line.startswith("F"):
                coordinates = re.split("(?<! ) ", line)
                for coordinate in coordinates:
                    if x is None:
                        x = float(coordinate.strip()) * BOHR_TO_ANG
                    elif y is None:
                        y = float(coordinate.strip()) * BOHR_TO_ANG
                    else:
                        molecule[atom_index].set_coordinates(np.array([x, y, float(coordinate.strip()) * BOHR_TO_ANG]))
                        x = None
                        y = None
                        atom_index += 1
                line = file.readline().strip()
        elif "NBond" in line.strip():
            line = file.readline().strip()
            number_of_bonds = [0]
            while "IBond" not in line:
                for bonds in re.split("(?<! ) ", line):
                    number_of_bonds.append(int(bonds.strip()))
                line = file.readline().strip()
            bond_number = 0
            atom_index = 1
            line = file.readline().strip()
            while "RBond" not in line:
                for bond in re.split("(?<! ) ", line):
                    if int(bond.strip()) != 0:
                        while bond_number == number_of_bonds[atom_index]:
                            atom_index += 1
                            bond_number = 0
                        molecule[atom_index].add_bond(int(bond.strip()))
                        bond_number += 1
                line = file.readline().strip()
            # edge case for H2
            if len(molecule) == 3 and molecule[1].atomic_number == 1 and molecule[2].atomic_number == 1:
                molecule[1].add_bond(2)
                molecule[2].add_bond(1)
        elif "Cartesian Force Constants" in line.strip():
            line = file.readline()
            while "Dipole" not in line:
                force_constants += line
                line = file.readline()
        line = file.readline()
    return force_constants


def read_nesi(file):
    line = file.readline()
    while line != "":
        if "Output coordinates in angstroms" in line.strip():
            molecule.clear()
            for i in range(3):
                file.readline()
            line = file.readline().strip()
            while not line == "":
                atomic_parameters = re.split("(?<! ) ", line)
                for i in range(len(atomic_parameters)):
                    atomic_parameters[i] = atomic_parameters[i].strip()
                if len(molecule) == 0:
                    # blank atom to align indexes with input values
                    molecule.append(Atom())
                molecule.append(Atom(int(round(float(atomic_parameters[2]))), np.array(
                    [float(atomic_parameters[3]), float(atomic_parameters[4]), float(atomic_parameters[5])])))
                line = file.readline().strip()
        elif "Z-matrix" in line.strip() and len(molecule) != 0:
            for atom in molecule:
                atom.clear_bonds()
            for i in range(6):
                file.readline()
            line = file.readline().strip()
            while "Stretch" in line:
                bonded_parameters = re.split("(?<! ) ", line)
                atom_1 = int(bonded_parameters[2].strip())
                atom_2 = int(bonded_parameters[3].strip())
                molecule[atom_1].add_bond(atom_2)
                molecule[atom_2].add_bond(atom_1)
                line = file.readline().strip()
        line = file.readline()
    return ""


def print_output(force_constants):

    output_file = filedialog.asksaveasfile(mode='w', defaultextension=".dat")
    if output_file is None:
        return
    ffc_file = open(output_file.name.strip(".dat") + ".ffc", "w")
    ffc_file.write(force_constants)
    ffc_file.close()

    number_of_int_coord = len(molecule_bonds) + len(angles) + len(linear_bends) + len(torsions) + len(plane_bends)
    number_of_available_motions = 5
    for angle in angles:
        if calculate_angle(molecule[angle.atom_1], molecule[angle.atom_2], molecule[angle.atom_3]) != 180.0:
            number_of_available_motions = 6

    output_text = ";options\n;calculations using internal coordinates\n;rotational constants are calculated in " \
                  "MHz\nform\n-4\nrothz\n;config\nscale\n{0}\n{0}*1.\nlocal\nfreq\n5.\ndata\nTitle Card " \
                  "Required\n'C1' 0 {0} 0 0\n{1} {0} 0 0 {2}\n{3} {4} {5} {6} {7}\n0 {0}*1 0\n\n\n"\
        .format(number_of_int_coord, len(molecule) - 1, number_of_available_motions, len(molecule_bonds), len(angles),
                len(plane_bends), len(linear_bends), len(torsions))

    for atom in molecule[1:]:
        output_text += "'{}{}'    {}    {}    {}    {}\n"\
            .format(atom.atomic_symbol, molecule.index(atom), atom.atomic_mass,
                    atom.coordinates[0], atom.coordinates[1], atom.coordinates[2])
    output_text += "\n\n"

    for bond in molecule_bonds:
        output_text += bond.__repr__() + "\n"
    if len(molecule_bonds) > 0:
        output_text += "\n\n"
    for angle in angles:
        output_text += angle.__repr__() + "\n"
    if len(angles) > 0:
        output_text += "\n\n"
    for plane_bend in plane_bends:
        output_text += plane_bend.__repr__() + "\n"
    if len(plane_bends) > 0:
        output_text += "\n\n"
    for linear_bend in linear_bends:
        output_text += linear_bend.__repr__() + "\n"
    if len(linear_bends) > 0:
        output_text += "\n\n"
    for torsion in torsions:
        output_text += torsion.__repr__() + "\n"
    if len(torsions) > 0:
        output_text += "\n\n"

    output_text += "376.00\n\n\n{} {}\n\n\n"\
        .format(int((len(molecule) - 1) * (len(molecule) - 2) / 2), len(molecule_bonds))

    for bond in molecule_bonds:
        output_text += bond.__repr__() + "\n"
    for i in range(1, len(molecule)):
        for j in range(1, len(molecule)):
            distance = Bond(i, j)
            if i != j and distance not in molecule_bonds:
                output_text += distance.__repr__() + "\n" \
                                                     ""
    output_file.write(output_text)
    output_file.close()


# could change to be recursive
def calculate_int_coord():
    for atom_1 in range(1, len(molecule)):
        for atom_2 in molecule[atom_1].bonds:
            for atom_3 in molecule[atom_2].bonds:
                if atom_3 != atom_1:
                    for atom_4 in molecule[atom_2].bonds:
                        if atom_4 != atom_1 and atom_4 != atom_3 and is_planar(molecule[atom_1], molecule[atom_2], molecule[atom_3], molecule[atom_4]):
                            bonded_atoms = [atom_1, atom_3, atom_4]
                            bonded_atoms.sort()
                            plane_bends.add(PlaneBend(bonded_atoms[0], bonded_atoms[1], bonded_atoms[2], atom_2))
                    if round(calculate_angle(molecule[atom_1], molecule[atom_2], molecule[atom_3])) == 180:
                        linear_bends.add(Angle(atom_1, atom_2, atom_3))
                        for atom_4 in molecule[atom_3].bonds:
                            if atom_4 != atom_2 and atom_4 != atom_1:
                                plane_bends.add(PlaneBend(atom_3, atom_2, atom_1, atom_4))
                    else:
                        angles.add(Angle(atom_1, atom_2, atom_3))
                        for atom_4 in molecule[atom_3].bonds:
                            if atom_4 != atom_2 and atom_4 != atom_1:
                                if round(calculate_angle(molecule[atom_2], molecule[atom_3], molecule[atom_4])) == 180:
                                    plane_bends.add(PlaneBend(atom_2, atom_3, atom_4, atom_1))
                                else:
                                    a_bonded = [atom_1]
                                    for a_bond in molecule[atom_2].bonds:
                                        if a_bond != atom_1 and a_bond != atom_3 and round(calculate_angle(molecule[a_bond], molecule[atom_2], molecule[atom_3])) != 180:
                                            a_bonded.append(a_bond)
                                    b_bonded = [atom_4]
                                    for b_bond in molecule[atom_3].bonds:
                                        if b_bond != atom_2 and b_bond != atom_4 and round(calculate_angle(molecule[b_bond], molecule[atom_3], molecule[atom_2])) != 180:
                                            b_bonded.append(b_bond)
                                    torsions.add(Torsion(atom_2, atom_3, a_bonded, b_bonded))


class Atom:
    def __init__(self, atomic_number=0, coordinates=np.array([0.0, 0.0, 0.0])):
        self.atomic_number = atomic_number
        self.atomic_symbol = atom_dict[atomic_number][0]
        self.atomic_mass = atom_dict[atomic_number][1]
        self.coordinates = coordinates
        self.bonds = set()

    def add_bond(self, atom_uid):
        self.bonds.add(atom_uid)
        molecule_bonds.add(Bond(molecule.index(self), atom_uid))

    def clear_bonds(self):
        self.bonds.clear()

    def set_coordinates(self, coordinates):
        self.coordinates = coordinates


class Bond:
    def __init__(self, atom_1, atom_2):
        self.atom_1 = atom_1
        self.atom_2 = atom_2

    def __repr__(self):
        return "{} {}".format(self.atom_1, self.atom_2) if self.atom_1 < self.atom_2 \
            else "{} {}".format(self.atom_2, self.atom_1)

    def __eq__(self, other):
        if isinstance(other, Bond):
            return self.__repr__() == other.__repr__()
        else:
            return False

    def __hash__(self):
        return hash(self.__repr__())


class Angle:
    def __init__(self, atom_1, atom_2, atom_3):
        self.atom_1 = atom_1
        self.atom_2 = atom_2
        self.atom_3 = atom_3

    def __repr__(self):
        return "{} {} {}".format(self.atom_1, self.atom_2, self.atom_3) if self.atom_1 < self.atom_3 \
            else "{} {} {}".format(self.atom_3, self.atom_2, self.atom_1)

    def __eq__(self, other):
        if isinstance(other, Angle):
            return self.__repr__() == other.__repr__()
        else:
            return False

    def __hash__(self):
        return hash(self.__repr__())


class Torsion:
    def __init__(self, atom_a, atom_b, a_atoms, b_atoms):
        self.atom_a = atom_a
        self.atom_b = atom_b
        self.a_atoms = a_atoms
        self.b_atoms = b_atoms
        self.a_atoms.sort()
        self.b_atoms.sort()
        while len(self.a_atoms) != 3:
            self.a_atoms.append(0)
        while len(self.b_atoms) != 3:
            self.b_atoms.append(0)

    def __repr__(self):
        base_string = "{} {} {} {} {} {} {} {}"
        if self.atom_a < self.atom_b:
            return base_string.format(self.a_atoms[0], self.a_atoms[1], self.a_atoms[2], self.atom_a, self.atom_b, self.b_atoms[0], self.b_atoms[1], self.b_atoms[2])
        else:
            return base_string.format(self.b_atoms[0], self.b_atoms[1], self.b_atoms[2], self.atom_b, self.atom_a, self.a_atoms[0], self.a_atoms[1], self.a_atoms[2])

    def __eq__(self, other):
        if isinstance(other, Torsion):
            return self.__repr__() == other.__repr__()
        else:
            return False

    def __hash__(self):
        return hash(self.__repr__())


class PlaneBend:
    def __init__(self, atom_1, atom_2, atom_3, reference_atom):
        self.atom_1 = atom_1
        self.atom_2 = atom_2
        self.atom_3 = atom_3
        self.reference_atom = reference_atom

    def __repr__(self):
        return "{} {} {} {}".format(self.atom_1, self.atom_2, self.atom_3, self.reference_atom)

    def __eq__(self, other):
        if isinstance(other, PlaneBend):
            return self.__repr__()[:5] == other.__repr__()[:5]
        else:
            return False

    def __hash__(self):
        return hash(self.__repr__())


def calculate_distance(atom_1, atom_2):
    return np.linalg.norm(atom_1.coordinates - atom_2.coordinates)


def calculate_angle(atom_1, atom_2, atom_3):

    vector_1 = atom_1.coordinates - atom_2.coordinates
    vector_2 = atom_3.coordinates - atom_2.coordinates

    cosine_angle = np.dot(vector_1, vector_2) / (np.linalg.norm(vector_1) * np.linalg.norm(vector_2))
    angle = np.arccos(cosine_angle)
    return float(np.degrees(angle))


def calculate_dihedral_angle(atom_1, atom_2, atom_3, atom_4):

    vector_1 = -1.0*(atom_2.coordinates - atom_1.coordinates)
    vector_2 = atom_3.coordinates - atom_2.coordinates
    vector_3 = atom_4.coordinates - atom_3.coordinates

    cross_1 = np.cross(vector_1, vector_2)
    cross_2 = np.cross(vector_3, vector_2)

    cross_3 = np.cross(cross_1, cross_2)

    dot_1 = np.dot(cross_3, vector_2)*(1.0/np.linalg.norm(vector_2))
    dot_2 = np.dot(cross_1, cross_2)

    return np.degrees(np.arctan2(dot_2, dot_1))


def is_planar(atom_1, atom_2, atom_3, atom_4):

    vector_1 = atom_2.coordinates - atom_1.coordinates
    vector_2 = atom_3.coordinates - atom_1.coordinates

    cross_1 = np.cross(vector_1, vector_2)

    return round(float(np.dot(cross_1, atom_1.coordinates)), 1) == round(float(np.dot(cross_1, atom_4.coordinates)), 1)


def select_file():
    filename = filedialog.askopenfilename(initialdir="/", title="Select file", filetypes=(
        ("all files", "*.*"), ("Output files", "*.out"), ("Checkpoint files", "*.fchk*")))
    if not (filename.endswith(".out") or filename.endswith(".fchk")):
        if filename != "":
            mbox.showerror("Error", "File type not supported")
    else:
        force_constants = identify_input(filename)
        calculate_int_coord()
        print_output(force_constants)
        current_status.set("File Processed. Awaiting further files")


root = Tk()
root.title("Shrink Preprocessor")
root.geometry("450x80")
current_status = StringVar()
current_status.set("Awaiting Gaussian output, Gaussian formatted checkpoint, or NWChem output file")
status_text = Label(root, textvariable=current_status)
status_text.grid(row=0, column=0)
upload_button = Button(root, text="Select input file", command=select_file)
upload_button.grid(row=1, column=0)
root.grid_rowconfigure(0, weight=1)
root.grid_rowconfigure(1, weight=1)
root.grid_columnconfigure(0, weight=1)

atom_dict = {
    0: ("", 0.0),
    1: ("H", 1.00783),
    2: ("He", 4.0026),
    3: ("Li", 7.016),
    4: ("Be", 9.01218),
    5: ("B", 11.00931),
    6: ("C", 12.0),
    7: ("N", 14.00307),
    8: ("O", 15.99491),
    9: ("F", 18.9984),
    10: ("Ne", 19.99244),
    11: ("Na", 22.9898),
    12: ("Mg", 23.98504),
    13: ("Al", 26.98154),
    14: ("Si", 27.97693),
    15: ("P", 30.97376),
    16: ("S", 31.97207),
    17: ("Cl", 34.96885),
    18: ("Ar", 39.9624),
    19: ("K", 38.96371),
    20: ("Ca", 39.96259),
    21: ("Sc", 44.95592),
    22: ("Ti", 45.948),
    23: ("V", 50.944),
    24: ("Cr", 51.9405),
    25: ("Mn", 54.9381),
    26: ("Fe", 55.9349),
    27: ("Co", 58.9332),
    28: ("Ni", 57.9353),
    29: ("Cu", 62.9298),
    30: ("Zn", 63.9291),
    31: ("Ga", 68.9257),
    32: ("Ge", 73.9219),
    33: ("As", 74.9216),
    34: ("Se", 79.9165),
    35: ("Br", 78.9183),
    36: ("Kr", 83.912),
    37: ("Rb", 84.9117),
    38: ("Sr", 87.9056),
    39: ("Y", 88.9054),
    40: ("Zr", 89.9043),
    41: ("Nb", 92.906),
    42: ("Mo", 97.9055),
    43: ("Tc", 97.9072),
    44: ("Ru", 101.9037),
    45: ("Rh", 102.9048),
    46: ("Pd", 105.9032),
    47: ("Ag", 106.90509),
    48: ("Cd", 113.9036),
    49: ("In", 114.9041),
    50: ("Sn", 117.9018),
    51: ("Sb", 120.9038),
    52: ("Te", 129.9067),
    53: ("I", 126.9004),
    54: ("Xe", 131.9042),
    55: ("Cs", 132.9051),
    56: ("Ba", 137.905),
    57: ("La", 138.9061),
    58: ("Ce", 139.9053),
    59: ("Pr", 140.9074),
    60: ("Nd", 143.9099),
    61: ("Pm", 144.9128),
    62: ("Sm", 151.9195),
    63: ("Eu", 152.9209),
    64: ("Gd", 157.9241),
    65: ("Tb", 159.925),
    66: ("Dy", 163.9288),
    67: ("Ho", 164.9303),
    68: ("Er", 165.9304),
    69: ("Tm", 168.9344),
    70: ("Yb", 173.939),
    71: ("Lu", 174.9409),
    72: ("Hf", 179.9468),
    73: ("Ta", 180.948),
    74: ("W", 183.951),
    75: ("Re", 186.956),
    76: ("Os", 189.9586),
    77: ("Ir", 192.9633),
    78: ("Pt", 194.9648),
    79: ("Au", 196.9666),
    80: ("Hg", 201.9706),
    81: ("Tl", 204.9745),
    82: ("Pb", 207.9766),
    83: ("Bi", 208.9804),
    84: ("Po", 208.9825),
    85: ("At", 210.9875),
    86: ("Rn", 222.0175),
    87: ("Fr", 223.0198),
    88: ("Ra", 226.0254),
    89: ("Ac", 227.0278),
    90: ("Th", 232.0382),
    91: ("Pa", 231.0359),
    92: ("U", 238.0508),
    93: ("Np", 237.048),
    94: ("Pu", 244.0642),
    95: ("Am", 243.0614),
    96: ("Cm", 247.0704),
    97: ("Bk", 247.0702),
    98: ("Cf", 249.0748),
    99: ("Es", 254.0881),
    100: ("Fm", 0.0),
    101: ("Md", 0.0),
    102: ("No", 0.0),
    103: ("Lr", 0.0),
    104: ("X", 0.0)
}

root.mainloop()
