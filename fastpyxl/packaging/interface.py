# Copyright (c) 2010-2024 fastpyxl

from abc import abstractmethod
from fastpyxl.compat.abc import ABC


class ISerialisableFile(ABC):

    """
    Interface for Serialisable classes that represent files in the archive
    """


    @property
    @abstractmethod
    def id(self):
        """
        Object id making it unique
        """
        pass


    @property
    @abstractmethod
    def _path(self):
        """
        File path in the archive
        """
        pass


    @property
    @abstractmethod
    def _namespace(self):
        """
        Qualified namespace when serialised
        """
        pass


    @property
    @abstractmethod
    def _type(self):
        """
        The content type for the manifest
        """


    @property
    @abstractmethod
    def _rel_type(self):
        """
        The content type for relationships
        """


    @property
    @abstractmethod
    def _rel_id(self):
        """
        Links object with parent
        """
