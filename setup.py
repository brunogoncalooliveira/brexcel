from setuptools import setup


setup(name='brexcel',
      version='0.2',
      description='Converts excel to a python dict and vice-versa',
      classifiers=[
        'Development Status :: 2 - Pre-Alpha',
        'License :: OSI Approved :: MIT License',
        'Programming Language :: Python :: 2.7',
        'Topic :: Text Processing :: Linguistic',
      ],
      keywords='excel dict convert xls xlsx',
      url='https://github.com/brunogoncalooliveira/brexcel',
      author='Bruno Oliveira',
      license='MIT',
      packages=['brexcel'],
      install_requires=['openpyxl'],
      include_package_data=True,
      zip_safe=False)