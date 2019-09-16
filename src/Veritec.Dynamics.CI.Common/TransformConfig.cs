using System.Collections.Generic;

namespace Veritec.Dynamics.CI.Common
{
    public class TransformConfig
    {
        public TransformConfig(List<Transform> transforms)
        {
            Transforms = transforms;
        }
        public Transform this[string entity, string attribute, string value]
        {
            get
            {
                if (ContainsKey(entity, attribute, value))
                    return Transforms.Find(x => x.TargetEntity == entity && x.TargetAttribute == attribute && x.TargetValue.ToLower() == value.ToLower());
                return null;
            }
        }

        public Transform this[string entity, string attribute]
        {
            get
            {
                if (ContainsKey(entity, attribute))
                    return Transforms.Find(x => x.TargetEntity == entity && x.TargetAttribute == attribute);
                return null;
            }
        }

        public List<Transform> Transforms { get; set; }

        public bool ContainsKey(string entity, string attribute, string value)
        {
            if (value == null) return false;

            var li = Transforms.FindAll(x => x.TargetEntity == entity && x.TargetAttribute == attribute && x.TargetValue.ToLower() == value.ToLower());

            if (li.Count > 0)
                return true;
            return false;
        }

        public bool ContainsKey(string entity, string attribute)
        {
            if (entity == null || attribute == null) return false;

            var li = Transforms.FindAll(x => x.TargetEntity == entity && x.TargetAttribute == attribute);

            if (li.Count > 0)
                return true;
            return false;
        }
    }
    public class Transform
    {
        public string TargetEntity { get; set; }
        public string TargetAttribute { get; set; }
        public string TargetValue { get; set; }
        public string ReplacementValue { get; set; }
        public string ReplacementAttribute { get; set; }

        public Transform(string targetEntity, string targetAttribute, string targetValue, string replacementValue, string replacementAttribute = null)
        {
            TargetEntity = targetEntity;
            TargetAttribute = targetAttribute;
            TargetValue = targetValue;
            ReplacementValue = replacementValue;

            // if the user doesn't specify a target attribute, the just 
            // target the source attribute
            ReplacementAttribute = replacementAttribute ?? targetAttribute;
        }
    }
}
