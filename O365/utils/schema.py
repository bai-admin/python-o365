from __future__ import annotations

import io
from typing import Any, ClassVar

try:
    import xmlschema
except ImportError:
    xmlschema = None


class Schema:
    """Class representing the OData metadata for Microsoft Graph API.

    This class provides methods to explore the Graph API schema, discover entity types,
    and understand relationships between resources in the API.
    """

    METADATA_URL: ClassVar[str] = "https://graph.microsoft.com/v1.0/$metadata"

    def __init__(self, parent: Any = None, **kwargs: Any) -> None:
        """Initialize Schema object for exploring Microsoft Graph API.

        :param parent: Parent object (Connection)
        :param kwargs: Additional keyword arguments
        """
        self.parent = parent
        self._parsed_schema = None

        # Load metadata during initialization
        self.parse_metadata()

    def get_metadata(self) -> str | None:
        """Get the metadata document from the Graph API.

        :return: Metadata document as a string or None if not available
        """
        if self.parent is None:
            return None

        response = self.parent.get(self.METADATA_URL)
        if not response:
            return None

        return response.text

    def parse_metadata(self) -> dict[str, Any] | None:
        """Parse the metadata document using xmlschema.

        :return: Parsed metadata as a dictionary or None if parsing failed
        """
        if not xmlschema:
            print("The xmlschema package is required for schema parsing.")
            print("Please install it with: pip install xmlschema")
            return None

        metadata = self.get_metadata()
        if not metadata:
            return None

        try:
            xs = xmlschema.XMLSchema(io.StringIO(metadata))
            xml_dict = xs.to_dict(io.StringIO(metadata))
            self._parsed_schema = xml_dict
            return xml_dict
        except Exception as e:
            print(f"Error parsing metadata: {e!s}")
            return None

    def get_entity_types(self) -> dict[str, Any] | None:
        """Return all entity types defined in the schema.

        :return: Dictionary of entity types or None if schema not parsed
        """
        if not self._parsed_schema:
            self.parse_metadata()

        if not self._parsed_schema:
            return None

        entity_types = {}
        data_services = self._parsed_schema.get("DataServices", {})
        schemas = data_services.get("Schema", [])

        if not isinstance(schemas, list):
            schemas = [schemas]

        for schema in schemas:
            namespace = schema.get("@Namespace")
            entities = schema.get("EntityType", [])

            if not isinstance(entities, list):
                entities = [entities]

            for entity in entities:
                name = entity.get("@Name")
                if name:
                    key = f"{namespace}.{name}" if namespace else name
                    entity_types[key] = entity

        return entity_types

    def get_complex_types(self) -> dict[str, Any] | None:
        """Return all complex types defined in the schema.

        :return: Dictionary of complex types or None if schema not parsed
        """
        if not self._parsed_schema:
            self.parse_metadata()

        if not self._parsed_schema:
            return None

        complex_types = {}
        data_services = self._parsed_schema.get("DataServices", {})
        schemas = data_services.get("Schema", [])

        if not isinstance(schemas, list):
            schemas = [schemas]

        for schema in schemas:
            namespace = schema.get("@Namespace")
            complex_list = schema.get("ComplexType", [])

            if not isinstance(complex_list, list):
                complex_list = [complex_list]

            for complex_type in complex_list:
                name = complex_type.get("@Name")
                if name:
                    key = f"{namespace}.{name}" if namespace else name
                    complex_types[key] = complex_type

        return complex_types

    def get_enums(self) -> dict[str, Any] | None:
        """Return all enumeration types defined in the schema.

        :return: Dictionary of enumeration types or None if schema not parsed
        """
        if not self._parsed_schema:
            self.parse_metadata()

        if not self._parsed_schema:
            return None

        enums = {}
        data_services = self._parsed_schema.get("DataServices", {})
        schemas = data_services.get("Schema", [])

        if not isinstance(schemas, list):
            schemas = [schemas]

        for schema in schemas:
            namespace = schema.get("@Namespace")
            enums_list = schema.get("EnumType", [])

            if not isinstance(enums_list, list):
                enums_list = [enums_list]

            for enum_type in enums_list:
                name = enum_type.get("@Name")
                if name:
                    key = f"{namespace}.{name}" if namespace else name
                    enum_type["members"] = self._process_enum_members(enum_type)
                    enums[key] = enum_type

        return enums

    def _process_enum_members(self, enum_type: dict[str, Any]) -> dict[str, int]:
        """Process enum members into a more usable format.

        :param enum_type: Enum type definition from the schema
        :return: Dictionary mapping member names to their values
        """
        members = {}
        enum_members = enum_type.get("Member", [])

        if not isinstance(enum_members, list):
            enum_members = [enum_members]

        for member in enum_members:
            name = member.get("@Name")
            value = member.get("@Value")
            if name and value:
                members[name] = int(value)

        return members

    def get_entity_properties(
        self,
        entity_name: str,
        namespace: str = "microsoft.graph",
    ) -> dict[str, Any] | None:
        """Return the properties of a specific entity.

        :param entity_name: Name of the entity
        :param namespace: Namespace of the entity (default: microsoft.graph)
        :return: Dictionary of properties and navigation properties
        """
        entity_types = self.get_entity_types()
        if not entity_types:
            return None

        full_name = f"{namespace}.{entity_name}"
        entity = entity_types.get(full_name)
        if not entity:
            return None

        # Build hierarchy of properties, including inherited ones
        properties = {}
        nav_properties = {}

        # Get base type and check for inheritance
        base_type = entity.get("@BaseType")
        if base_type:
            # Get properties from the base type
            base_properties = self.get_entity_properties(
                base_type.split(".")[-1],
                ".".join(base_type.split(".")[:-1]),
            )
            if base_properties:
                properties.update(base_properties.get("properties", {}))
                nav_properties.update(base_properties.get("navigation_properties", {}))

        # Get entity's own properties
        entity_properties = entity.get("Property", [])
        if not isinstance(entity_properties, list):
            entity_properties = [entity_properties]

        for prop in entity_properties:
            name = prop.get("@Name")
            if name:
                properties[name] = prop

        # Get entity's own navigation properties
        entity_nav_props = entity.get("NavigationProperty", [])
        if not isinstance(entity_nav_props, list):
            entity_nav_props = [entity_nav_props]

        for nav_prop in entity_nav_props:
            name = nav_prop.get("@Name")
            if name:
                nav_properties[name] = nav_prop

        return {
            "properties": properties,
            "navigation_properties": nav_properties,
        }

    def get_entity_inheritance(
        self,
        entity_name: str,
        namespace: str = "microsoft.graph",
    ) -> list[str] | None:
        """Return the inheritance chain for an entity.

        :param entity_name: Name of the entity
        :param namespace: Namespace of the entity (default: microsoft.graph)
        :return: List of base types in inheritance chain
        """
        entity_types = self.get_entity_types()
        if not entity_types:
            return None

        full_name = f"{namespace}.{entity_name}"
        entity = entity_types.get(full_name)
        if not entity:
            return None

        inheritance = []
        current_entity = entity
        current_name = full_name

        while current_entity:
            base_type = current_entity.get("@BaseType")
            if not base_type:
                break

            inheritance.append(base_type)

            # Get the base entity
            base_entity = entity_types.get(base_type)
            if not base_entity:
                break

            current_entity = base_entity
            current_name = base_type

        return inheritance

    def find_resource_definition(
        self,
        resource_name: str,
        namespace: str = "microsoft.graph",
    ) -> dict[str, Any] | None:
        """Find and return the definition of a resource type in the schema.

        :param resource_name: Name of the resource to find
        :param namespace: Namespace of the resource (default: microsoft.graph)
        :return: Dictionary with resource definition and type
        """
        if not self._parsed_schema:
            self.parse_metadata()

        if not self._parsed_schema:
            return None

        full_name = f"{namespace}.{resource_name}"

        # Check entity types
        entity_types = self.get_entity_types() or {}
        if full_name in entity_types:
            return {"type": "EntityType", "definition": entity_types[full_name]}

        # Check complex types
        complex_types = self.get_complex_types() or {}
        if full_name in complex_types:
            return {"type": "ComplexType", "definition": complex_types[full_name]}

        # Check enum types
        enum_types = self.get_enums() or {}
        if full_name in enum_types:
            return {"type": "EnumType", "definition": enum_types[full_name]}

        return None

    def is_collection_type(self, type_name: str) -> bool:
        """Check if a type is a collection.

        :param type_name: Name of the type
        :return: True if it's a collection type, False otherwise
        """
        if not type_name:
            return False

        return type_name.startswith("Collection(")

    def get_base_type(self, type_name: str) -> str:
        """Extract the base type from a collection type.

        :param type_name: Name of the type
        :return: Base type without the Collection wrapper
        """
        if not self.is_collection_type(type_name):
            return type_name

        # Extract the type inside Collection(...)
        return type_name[11:-1]  # Remove 'Collection(' and ')'

    def get_property_type_info(
        self,
        property_def: dict[str, Any],
    ) -> dict[str, Any]:
        """Get detailed information about a property's type.

        :param property_def: Property definition from the schema
        :return: Dictionary with type information
        """
        type_info = {
            "name": property_def.get("@Name", ""),
            "type": property_def.get("@Type", ""),
            "nullable": property_def.get("@Nullable", "false") == "true",
        }

        # Check if it's a collection
        is_collection = self.is_collection_type(type_info["type"])
        type_info["is_collection"] = is_collection

        # Get the base type for collections
        if is_collection:
            base_type = self.get_base_type(type_info["type"])
            type_info["base_type"] = base_type

            # Look up definition of the base type
            if not base_type.startswith("Edm."):
                base_parts = base_type.split(".")
                if len(base_parts) > 1:
                    namespace = ".".join(base_parts[:-1])
                    resource_name = base_parts[-1]
                    type_info["resource_def"] = self.find_resource_definition(
                        resource_name,
                        namespace,
                    )
        elif not type_info["type"].startswith("Edm."):
            # For non-collection non-primitive types, look up their definition
            type_parts = type_info["type"].split(".")
            if len(type_parts) > 1:
                namespace = ".".join(type_parts[:-1])
                resource_name = type_parts[-1]
                type_info["resource_def"] = self.find_resource_definition(
                    resource_name,
                    namespace,
                )

        return type_info

    def get_navigation_property_type_info(
        self,
        nav_property_def: dict[str, Any],
    ) -> dict[str, Any]:
        """Get detailed information about a navigation property's type.

        :param nav_property_def: Navigation property definition from the schema
        :return: Dictionary with type information
        """
        type_info = {
            "name": nav_property_def.get("@Name", ""),
            "type": nav_property_def.get("@Type", ""),
            "nullable": nav_property_def.get("@Nullable", "false") == "true",
            "contains_target": nav_property_def.get("@ContainsTarget", "false")
            == "true",
        }

        # Check if it's a collection
        is_collection = self.is_collection_type(type_info["type"])
        type_info["is_collection"] = is_collection

        # Get the base type for collections
        if is_collection:
            base_type = self.get_base_type(type_info["type"])
            type_info["base_type"] = base_type

            # Look up definition of the base type
            base_parts = base_type.split(".")
            if len(base_parts) > 1:
                namespace = ".".join(base_parts[:-1])
                resource_name = base_parts[-1]
                type_info["resource_def"] = self.find_resource_definition(
                    resource_name,
                    namespace,
                )
        else:
            # For non-collection types, look up their definition
            type_parts = type_info["type"].split(".")
            if len(type_parts) > 1:
                namespace = ".".join(type_parts[:-1])
                resource_name = type_parts[-1]
                type_info["resource_def"] = self.find_resource_definition(
                    resource_name,
                    namespace,
                )

        return type_info

    def get_resource_structure(
        self,
        resource_name: str,
        namespace: str = "microsoft.graph",
        max_depth: int = 5,
        visited: set[str] = None,
    ) -> dict[str, Any] | None:
        """Recursively explore a resource type and build a complete structure representation.

        This method navigates through the resource type definition and builds a hierarchical
        representation of its properties and navigation properties. It follows relationships
        to other resource types up to the specified maximum depth to prevent infinite recursion.

        :param resource_name: Name of the resource to explore
        :param namespace: Namespace of the resource (default: microsoft.graph)
        :param max_depth: Maximum recursion depth to prevent infinite loops
        :param visited: Set of already visited resource types
        :return: Dictionary with the complete resource structure
        """
        if max_depth <= 0:
            return {"_max_depth_reached": True}

        if visited is None:
            visited = set()

        full_name = f"{namespace}.{resource_name}"

        # Check if we've already visited this resource type
        if full_name in visited:
            return {"_type": full_name, "_cycle_detected": True}

        # Add this resource to the visited set
        visited.add(full_name)

        # Find the resource definition
        resource_def = self.find_resource_definition(resource_name, namespace)
        if not resource_def:
            return None

        result = {
            "_type": full_name,
            "_resource_type": resource_def["type"],
        }

        # Process resource based on its type
        if resource_def["type"] in ("EntityType", "ComplexType"):
            entity = resource_def["definition"]

            # Get properties
            result["properties"] = {}
            properties = entity.get("Property", [])

            if not isinstance(properties, list):
                properties = [properties]

            for prop in properties:
                name = prop.get("@Name")
                if name:
                    prop_type_info = self.get_property_type_info(prop)
                    result["properties"][name] = prop_type_info

                    # Recursively process complex properties
                    if "resource_def" in prop_type_info and max_depth > 1:
                        if prop_type_info["resource_def"]["type"] in (
                            "ComplexType",
                            "EntityType",
                        ):
                            resource_name = prop_type_info["type"].split(".")[-1]
                            if self.is_collection_type(prop_type_info["type"]):
                                resource_name = self.get_base_type(
                                    prop_type_info["type"]
                                ).split(".")[-1]

                            namespace_parts = prop_type_info["type"].split(".")
                            namespace = ".".join(namespace_parts[:-1])
                            if self.is_collection_type(prop_type_info["type"]):
                                namespace_parts = self.get_base_type(
                                    prop_type_info["type"]
                                ).split(".")
                                namespace = ".".join(namespace_parts[:-1])

                            sub_result = self.get_resource_structure(
                                resource_name,
                                namespace,
                                max_depth - 1,
                                visited.copy(),
                            )

                            if sub_result:
                                prop_type_info["structure"] = sub_result

            # Get navigation properties (only for EntityType)
            if resource_def["type"] == "EntityType":
                result["navigation_properties"] = {}
                nav_properties = entity.get("NavigationProperty", [])

                if not isinstance(nav_properties, list):
                    nav_properties = [nav_properties]

                for nav_prop in nav_properties:
                    name = nav_prop.get("@Name")
                    if name:
                        nav_type_info = self.get_navigation_property_type_info(nav_prop)
                        result["navigation_properties"][name] = nav_type_info

                        # Recursively process navigation properties
                        if max_depth > 1:
                            resource_name = nav_type_info["type"].split(".")[-1]
                            if self.is_collection_type(nav_type_info["type"]):
                                resource_name = self.get_base_type(
                                    nav_type_info["type"]
                                ).split(".")[-1]

                            namespace_parts = nav_type_info["type"].split(".")
                            namespace = ".".join(namespace_parts[:-1])
                            if self.is_collection_type(nav_type_info["type"]):
                                namespace_parts = self.get_base_type(
                                    nav_type_info["type"]
                                ).split(".")
                                namespace = ".".join(namespace_parts[:-1])

                            sub_result = self.get_resource_structure(
                                resource_name,
                                namespace,
                                max_depth - 1,
                                visited.copy(),
                            )

                            if sub_result:
                                nav_type_info["structure"] = sub_result
        elif resource_def["type"] == "EnumType":
            # For enum types, add their members
            result["members"] = resource_def["definition"].get("members", {})

        return result
