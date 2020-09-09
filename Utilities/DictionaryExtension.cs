using System;
using System.Collections.Generic;
using System.Text;

namespace OfficeFireSync.Utilities
{
    public static class DictionaryExtension
    {
        public static bool DeepCompare<K, V>(this IDictionary<K, V> ob, IDictionary<K, V> ob2)
        {
            // Test for equality.
            bool equal = false;
            if (ob.Count == ob2.Count) // Require equal count.
            {
                equal = true;
                foreach (var pair in ob)
                {
                    if (ob2.TryGetValue(pair.Key, out V value))
                    {
                        // Require value be equal.
                        if (value.Equals(pair.Value))
                        {
                            equal = false;
                            break;
                        }
                    }
                    else
                    {
                        // Require key be present.
                        equal = false;
                        break;
                    }
                }
            }

            return equal;
        }
    }
}
