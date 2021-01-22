"""Points on Surface.
"""

__author__ = 'Jad Akkawi - jad.akkawi@dar.com // Ali Safeiddine - ali.safeiddine@dar.com'
__copyright__ = 'Autodesk 2020'
__version__ = '1.0.0'

import sys
sys.path.append('C:\\Program Files (x86)\\IronPython 2.7\\Lib')
from getpass import getuser
import math
import csv
import os
import clr
clr.AddReference("System.Windows.Forms")
clr.AddReference("mscorlib")
clr.AddReference("acdbmgdbrep")

from System import *
from System.IO import *
from System.Windows.Forms import *
from System.Runtime.InteropServices import *

from Autodesk.Civil.ApplicationServices import *
from Autodesk.Civil.DatabaseServices import *
from Autodesk.Civil import *

from Autodesk.AutoCAD.Runtime import *
from Autodesk.AutoCAD.ApplicationServices import *
import Autodesk.AutoCAD.ApplicationServices.Application as AcApp

from Autodesk.AutoCAD.Internal import *
from Autodesk.AutoCAD.BoundaryRepresentation import *

from Autodesk.AutoCAD.EditorInput import *
from Autodesk.AutoCAD.Geometry import *
from Autodesk.AutoCAD.Colors import *
from Autodesk.AutoCAD.DatabaseServices import *

from Autodesk.Aec.PropertyData import *
from Autodesk.Aec.PropertyData.DatabaseServices import *


def save_path():
    open_file_dialog = OpenFileDialog()
    open_file_dialog.InitialDirectory = "C:\\"
    open_file_dialog.Filter = "txt files (*.txt)|*.txt|All files (*.*)|*.*"
    open_file_dialog.FilterIndex = 2
    open_file_dialog.RestoreDirectory = True
    if open_file_dialog.ShowDialog() == DialogResult.OK:
        file_path = open_file_dialog.FileName
        return file_path


def is_inside(_pl1, _pl2):
    """
    Based on The Jordan Curve Theorem for Polygons.
    :param _pl1: supposedly the outer polyline
    :param _pl2: supposedly the inner polyline
    :return: bool
    """
    if _pl1.Area < _pl2.Area:
        return False
    else:
        vec = Vector3d(0, 0, 1)
        geom_pl1 = _pl1.GetGeCurve()
        horizontal_line = Line(Point3d(_pl2.StartPoint.X, _pl2.StartPoint.Y, _pl2.StartPoint.Z),
                             Point3d(_pl2.StartPoint.X + 1000000, _pl2.StartPoint.Y, _pl2.StartPoint.Z))
        geom_line = horizontal_line.GetGeCurve()
        inters = CurveCurveIntersector3d(geom_line, geom_pl1, vec)
        if inters.NumberOfIntersectionPoints % 2 == 0:
            return False
        else:
            tier_pt = _pl2.GetPointAtDist(_pl2.Length/3)
            horizontal_line = Line(Point3d(tier_pt.X, tier_pt.Y, tier_pt.Z),
                               Point3d(tier_pt.X + 1000000, tier_pt.Y, tier_pt.Z))
            geom_line = horizontal_line.GetGeCurve()
            inters = CurveCurveIntersector3d(geom_line, geom_pl1, vec)
            if inters.NumberOfIntersectionPoints % 2 == 0:
                    return False
            else:
                two_tier_pt = _pl2.GetPointAtDist(_pl2.Length * 2/3)
                horizontal_line = Line(Point3d(two_tier_pt.X, two_tier_pt.Y, two_tier_pt.Z),
                                   Point3d(two_tier_pt.X + 1000000, two_tier_pt.Y, two_tier_pt.Z))
                geom_line = horizontal_line.GetGeCurve()
                inters = CurveCurveIntersector3d(geom_line, geom_pl1, vec)
                if inters.NumberOfIntersectionPoints % 2 == 0:
                    return False
                else:
                    return True


def create_layer(_db, _tr, layer_name, aci_value):
    layer_table = _tr.GetObject(_db.LayerTableId, OpenMode.ForWrite)
    if not layer_table.Has(layer_name):
        new_layer_table_record = LayerTableRecord()
        new_layer_table_record.Color = Color.FromColorIndex(ColorMethod.ByAci, aci_value)
        new_layer_table_record.Name = layer_name
        layer_table.Add(new_layer_table_record)
        _tr.AddNewlyCreatedDBObject(new_layer_table_record, True)


def create_group(_db, _tr, _group_name):
    gd = _tr.GetObject(_db.GroupDictionaryId, OpenMode.ForRead)
    gd.UpgradeOpen()
    grp = Group(_group_name, True)
    gd.SetAt(_group_name, grp)
    _tr.AddNewlyCreatedDBObject(grp, True)
    return grp


def append_to_group(_obj, _group):
    _group.Append(_obj.Id)


def group_openings_with_slabs(_btr):
    layer_dict = {}
    for oid in _btr:
        dbo = oid.GetObject(OpenMode.ForWrite)
        if isinstance(dbo, Polyline):
            layer_dict.setdefault(str(dbo.Layer), [])
            layer_dict[str(dbo.Layer)].append([dbo, dbo.Area])
    for lay in layer_dict:
        layer_dict[lay].sort(key=lambda tup: tup[1], reverse=True)
        for pol in layer_dict[lay]:
            for poly in layer_dict[lay][layer_dict[lay].index(pol)+1:]:
                if is_inside(pol[0], poly[0]):
                    pol.append(poly[0])
                    layer_dict[lay].remove(poly)
            pol.pop(1)
        # MessageBox.Show(lay + str(layer_dict[lay]))
    return layer_dict


def create_and_populate_layers(_db, _tr, _layer_dict):
    create_layer(_db, _tr, "open_polyline", 1)
    for lay in _layer_dict:
        for i in range(len(_layer_dict[lay])):
            create_layer(_db, _tr, lay + " slab " + str(i) + " OUTER", 3)
            create_layer(_db, _tr, lay + " slab " + str(i) + " INNER", 3)
            for j in range(len(_layer_dict[lay][i])):
                if j == 0:
                    if _layer_dict[lay][i][j].Closed is True:
                        _layer_dict[lay][i][j].Layer = lay + " slab " + str(i) + " OUTER"
                    else:
                        _layer_dict[lay][i][j].Layer = "open_polyline"
                else:
                    if _layer_dict[lay][i][j].Closed is True:
                        _layer_dict[lay][i][j].Layer = lay + " slab " + str(i) + " INNER"
                    else:
                        _layer_dict[lay][i][j].Layer = "open_polyline"


def get_tin_surface(_ed, _tr):
    osr = PromptEntityOptions('\nSelect a TinSurface')
    osr.SetRejectMessage('\nObject must be a TinSurface.')
    osr.AddAllowedClass(TinSurface, False)
    sr = _ed.GetEntity(osr)
    srr = _tr.GetObject(sr.ObjectId, OpenMode.ForWrite)
    return srr


def get_width(_pll, _error_list):
    area = _pll.Area * 1.1
    length = _pll.Length
    width = None
    for i in range(5):
        try:
            area /= 1.1
            the_sum = float(length)/2
            delta = (the_sum**2) - 4 * area
            width = float((the_sum - math.sqrt(delta)))/2
            return width
        except ValueError:
            pass
    if width is None:
        width = 20
        _error_list.append(_pll)
        return width


def get_point_list_from_offset_curve(_pll, step, _error_list):
    if step == 0:
        geom_pll = _pll.GetGeCurve()
        return geom_pll.GetSamplePoints(max(_pll.NumberOfVertices * 2, int(float(_pll.Length)/50)))
    else:
        width = get_width(_pll, _error_list)
        if width > 15:
            offset_curve = _pll.GetOffsetCurves(- width * step / 10)
            if len(offset_curve) != 0:
                offsets = []
                for curve in offset_curve:
                    geom_offset = curve.GetGeCurve()
                    offsets.extend(geom_offset.GetSamplePoints(max(curve.NumberOfVertices * 2,
                                                           int(float(curve.Length)/50))))
                return offsets
            else:
                offset_curve = de_curve_poly(_pll).GetOffsetCurves(-width * step / 10)
                if len(offset_curve) != 0:
                    offsets = []
                    for curve in offset_curve:
                        geom_offset = curve.GetGeCurve()
                        offsets.extend(geom_offset.GetSamplePoints(max(curve.NumberOfVertices * 2,
                                                                       int(float(curve.Length)/50))))
                    return offsets
                elif step < 2.5:
                    _error_list.append([_pll])
                return []
        else:
            return []


def append_vertices(_pll, _list):
    for _index in range(_pll.NumberOfVertices-1):
        _list.append(_pll.GetPoint2dAt(_index))


def append_points(pnt_lst, _list):
    for point_on_curve in pnt_lst:
        _list.append(point_on_curve.Point)


def gather_points_for_slab(_slab, _error_list):
    p_list = []
    for i in range(len(_slab)):
        append_vertices(_slab[i], p_list)
        if i == 0:
            for j in range(5):
                append_points(get_point_list_from_offset_curve(_slab[i], j, _error_list), p_list)
        else:
            append_points(get_point_list_from_offset_curve(_slab[i], 0, _error_list), p_list)
    _slab.append(p_list)


def gather_points_for_all_slabs(_layer_dict):
    error_list = []
    for lay in _layer_dict:
        for slab in _layer_dict[lay]:
            gather_points_for_slab(slab, error_list)
    return error_list


def write_csv(_layer_dict, _graded_surface):

    with open('poly_points.csv', 'wb') as _file:
        writer = csv.writer(_file)
        for lay in _layer_dict:
            for i in range(len(_layer_dict[lay])):
                writer.writerow([lay + " slab " + str(i)])
                for point in _layer_dict[lay][i][-1]:
                    try:
                        z = _graded_surface.FindElevationAtXY(point.X, point.Y)
                        writer.writerow([point.X, point.Y, z])
                    except:
                        continue


def de_curve_poly(_pll):
    de_curved_poly = Polyline()
    vertices = []
    append_vertices(_pll, vertices)
    for i in range(len(vertices)):
        de_curved_poly.AddVertexAt(i, vertices[i], 0, 0, 0)
    return de_curved_poly


def main():
    """
        NOTE:   Takes Poly lines of slabs and openings, separates them by layers.

                Each layer gets a set of points projected to a surface by dividing poly lines into segments.

                Write points on csv by layer.
    """
    os.chdir("C:\\Users\\" + getuser() + "\\Desktop")
    try:
        doc = AcApp.DocumentManager.MdiActiveDocument
        with doc.LockDocument():
            with doc.Database as db:
                with db.TransactionManager.StartTransaction() as tr:

                    ed = doc.Editor
                    bt = tr.GetObject(db.BlockTableId, OpenMode.ForWrite)
                    btr = tr.GetObject(bt[BlockTableRecord.ModelSpace], OpenMode.ForWrite)

                    layer_dict = group_openings_with_slabs(btr)

                    create_and_populate_layers(db, tr, layer_dict)

                    error_list = gather_points_for_all_slabs(layer_dict)

                    graded_surface = get_tin_surface(ed, tr)

                    write_csv(layer_dict, graded_surface)

                    create_layer(db, tr, "width_error", 4)

                    create_layer(db, tr, "offset_error", 5)

                    create_layer(db, tr, "width_&_offset_error", 6)

                    create_layer(db, tr, "open_&_offset_error", 2)

                    create_layer(db, tr, "open_&_width_error", 7)

                    for err_poly in error_list:
                        if isinstance(err_poly, Polyline):
                            if err_poly.Layer != "open_polyline":
                                err_poly.Layer = "width_error"
                            else:
                                err_poly.Layer = "open_&_width_error"
                        if isinstance(err_poly, list):
                            if err_poly[0].Layer not in ["width_error", "open_polyline"]:
                                err_poly[0].Layer = "offset_error"
                            else:
                                if err_poly[0].Layer == "width_error":
                                    err_poly[0].Layer = "width_&_offset_error"
                                else:
                                    err_poly[0].Layer = "open_&_offset_error"
                    tr.Commit()
    except Exception() as ex:
        MessageBox.Show(ex.ToString())


main()

# watch out from using layers for errors it would ruin layer for import to revit
