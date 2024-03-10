#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""Script to generate Windows Shell link test files.

Requires Windows and pywin32.
"""

import os

import pythoncom

from win32com.propsys import propsys
from win32com.shell import shell


if __name__ == '__main__':
  # c98dce577f884ef8.customDestinations-ms
  destination_list = pythoncom.CoCreateInstance(
      shell.CLSID_DestinationList, None, pythoncom.CLSCTX_INPROC_SERVER,
      shell.IID_ICustomDestinationList)

  destination_list.SetAppID('empty')
  destination_list.BeginList()
  destination_list.CommitList()

  # e0ad294086e6f817.customDestinations-ms
  destination_list = pythoncom.CoCreateInstance(
      shell.CLSID_DestinationList, None, pythoncom.CLSCTX_INPROC_SERVER,
      shell.IID_ICustomDestinationList)

  object_collection = pythoncom.CoCreateInstance(
      shell.CLSID_EnumerableObjectCollection, None,
      pythoncom.CLSCTX_INPROC_SERVER, shell.IID_IObjectCollection)

  shortcut = pythoncom.CoCreateInstance(
      shell.CLSID_ShellLink, None, pythoncom.CLSCTX_INPROC_SERVER,
      shell.IID_IShellLink)

  shortcut.SetPath('C:\\test')
  shortcut.SetArguments('My Arguments')
  shortcut.SetIconLocation('My Icon', 0)

  property_store = shortcut.QueryInterface(propsys.IID_IPropertyStore)

  property_key = propsys.PSGetPropertyKeyFromName('System.Title')
  property_value = propsys.PROPVARIANTType('My Title', pythoncom.VT_BSTR)
  property_store.SetValue(property_key, property_value)

  object_collection.AddObject(shortcut)

  destination_list.SetAppID('category')
  destination_list.BeginList()
  destination_list.AppendCategory('My Category', object_collection)
  destination_list.CommitList()

  # a7d1f78372f80802.customDestinations-ms
  destination_list = pythoncom.CoCreateInstance(
      shell.CLSID_DestinationList, None, pythoncom.CLSCTX_INPROC_SERVER,
      shell.IID_ICustomDestinationList)

  destination_list.SetAppID('known_category1')
  destination_list.BeginList()
  destination_list.AppendKnownCategory(1)
  destination_list.CommitList()

  # 5bb38bf987d0997a.customDestinations-ms
  destination_list = pythoncom.CoCreateInstance(
      shell.CLSID_DestinationList, None, pythoncom.CLSCTX_INPROC_SERVER,
      shell.IID_ICustomDestinationList)

  destination_list.SetAppID('known_category2')
  destination_list.BeginList()
  destination_list.AppendKnownCategory(2)
  destination_list.CommitList()

  # f72557edb63b1aef.customDestinations-ms
  destination_list = pythoncom.CoCreateInstance(
      shell.CLSID_DestinationList, None, pythoncom.CLSCTX_INPROC_SERVER,
      shell.IID_ICustomDestinationList)

  object_collection = pythoncom.CoCreateInstance(
      shell.CLSID_EnumerableObjectCollection, None,
      pythoncom.CLSCTX_INPROC_SERVER, shell.IID_IObjectCollection)

  shortcut = pythoncom.CoCreateInstance(
      shell.CLSID_ShellLink, None, pythoncom.CLSCTX_INPROC_SERVER,
      shell.IID_IShellLink)

  shortcut.SetPath('C:\\test')
  shortcut.SetArguments('My Arguments')
  shortcut.SetIconLocation('My Icon', 0)

  property_store = shortcut.QueryInterface(propsys.IID_IPropertyStore)

  property_key = propsys.PSGetPropertyKeyFromName('System.Title')
  property_value = propsys.PROPVARIANTType('My Title', pythoncom.VT_BSTR)
  property_store.SetValue(property_key, property_value)

  object_collection.AddObject(shortcut)

  destination_list.SetAppID('user_tasks')
  destination_list.BeginList()
  destination_list.AddUserTasks(object_collection)
  destination_list.CommitList()

  # 368d807282ccde9d.customDestinations-ms
  destination_list = pythoncom.CoCreateInstance(
      shell.CLSID_DestinationList, None, pythoncom.CLSCTX_INPROC_SERVER,
      shell.IID_ICustomDestinationList)

  object_collection = pythoncom.CoCreateInstance(
      shell.CLSID_EnumerableObjectCollection, None,
      pythoncom.CLSCTX_INPROC_SERVER, shell.IID_IObjectCollection)

  shortcut = pythoncom.CoCreateInstance(
      shell.CLSID_ShellLink, None, pythoncom.CLSCTX_INPROC_SERVER,
      shell.IID_IShellLink)

  shortcut.SetPath('C:\\test')
  shortcut.SetArguments('My Arguments')
  shortcut.SetIconLocation('My Icon', 0)

  property_store = shortcut.QueryInterface(propsys.IID_IPropertyStore)

  property_key = propsys.PSGetPropertyKeyFromName('System.Title')
  property_value = propsys.PROPVARIANTType('My Title', pythoncom.VT_BSTR)
  property_store.SetValue(property_key, property_value)

  object_collection.AddObject(shortcut)

  destination_list.SetAppID('specimen')
  destination_list.BeginList()
  destination_list.AppendCategory('My Category 1', object_collection)
  destination_list.AppendKnownCategory(1)
  destination_list.AddUserTasks(object_collection)
  destination_list.AppendCategory('My Category 2', object_collection)
  destination_list.AppendKnownCategory(2)
  destination_list.CommitList()
