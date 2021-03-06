﻿using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml;
using System.Threading.Tasks;

namespace ConsoleApplication
{
    public abstract class SceneObjectDecorator : SceneObject
    {
        private SceneObject _sceneObject;

        public SceneObjectDecorator(SceneObject so)
        {
            _sceneObject = so;
        }

        public virtual object Clone()
        {
            return _sceneObject.Clone();
        }

        public virtual XmlElement getXMLTree()
        {
            return _sceneObject.getXMLTree();
        }

        public virtual void ConvertToYoobaUnits(int width, int height)
        {
            _sceneObject.ConvertToYoobaUnits(width, height);
        }

        public virtual Properties getProperties()
        {
            return _sceneObject.getProperties();
        }

        public virtual void setProperties(Properties properties)
        {
            _sceneObject.setProperties(properties);
        }

        public virtual void setXMLDocumentRoot(ref XmlDocument xmldocument)
        {
            _sceneObject.setXMLDocumentRoot(ref xmldocument);
        }

        public virtual void setObjectType(string objectType)
        {
            _sceneObject.setObjectType(objectType);
        }

        public virtual void setZindex(int z)
        {
            _sceneObject.setZindex(z);
        }

        public virtual XmlDocument getXMLDocumentRoot()
        {
            return _sceneObject.getXMLDocumentRoot();
        }
    }
}
